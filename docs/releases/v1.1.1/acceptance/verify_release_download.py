"""Verify a newly downloaded v1.1.1 ZIP against the frozen local build.

Usage: python verify_release_download.py DOWNLOAD.zip --output-dir NEW_DIRECTORY
No download is performed. The output directory must not already exist.
"""
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
MANIFEST = ROOT / "build" / "verification-manifest.json"
MEMBERS = ("ExcelTools.exe", "使用说明.txt")
TIMEOUT_SECONDS = 90


def sha(path):
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("zip_path", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    archive_path = args.zip_path.resolve(strict=True)
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {
        "recorded_utc": datetime.now(timezone.utc).isoformat(),
        "app_version": "1.1.1",
        "passed": False,
        "stage": "manifest",
        "script_sha256": sha(__file__),
        "timeout_seconds": TIMEOUT_SECONDS,
        "limits": "Local Windows self-test in isolated paths and environment; "
                  "not a clean-machine or business-workbook acceptance test.",
    }
    started = time.monotonic()
    try:
        require(os.name == "nt", "This release acceptance requires Windows")
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8-sig"))
        require(manifest["app_version"] == "1.1.1", "Unexpected manifest version")
        expected_zip = manifest["package"]["sha256"]
        expected_exe = manifest["executable"]["sha256"]
        report.update(
            manifest_sha256=sha(MANIFEST),
            expected_zip_sha256=expected_zip,
            expected_exe_sha256=expected_exe,
        )
        report["stage"] = "zip-hash"
        report["downloaded_zip_sha256"] = sha(archive_path)
        require(report["downloaded_zip_sha256"] == expected_zip,
                "Downloaded ZIP does not match frozen build")

        report["stage"] = "zip-members-and-crc"
        with ZipFile(archive_path) as archive:
            names = archive.namelist()
            require(sorted(names) == sorted(MEMBERS),
                    "ZIP must contain exactly ExcelTools.exe and 使用说明.txt")
            require(archive.testzip() is None, "ZIP CRC validation failed")
            report["members"] = [
                {"name": item.filename, "size_bytes": item.file_size,
                 "compressed_bytes": item.compress_size, "crc32": f"{item.CRC:08x}"}
                for item in archive.infolist()
            ]
            report["zip_crc_passed"] = True
            # Only this new ignored directory is used; existing releases are untouched.
            test_root = REPO / "testfile"
            test_root.mkdir(exist_ok=True)
            work = Path(tempfile.mkdtemp(prefix="v1.1.1-release-download-", dir=test_root))
            portable = work / "下载 解压验收"
            cwd = work / "独立 工作目录"
            extraction = work / "临时 解压目录"
            for folder in (portable, cwd, extraction):
                folder.mkdir()
            for name in MEMBERS:
                with archive.open(name) as source, (portable / name).open("xb") as target:
                    shutil.copyfileobj(source, target, length=1024 * 1024)
        report["work_directory_relative_to_repo"] = work.relative_to(REPO).as_posix()
        report["stage"] = "exe-hash"
        exe = portable / "ExcelTools.exe"
        report["extracted_exe_sha256"] = sha(exe)
        require(report["extracted_exe_sha256"] == expected_exe,
                "Extracted EXE does not match frozen build")

        system32 = Path(os.environ["SystemRoot"]) / "System32"
        env = {
            key: value for key, value in os.environ.items()
            if not key.upper().startswith(("PYTHON", "TCL", "TK", "CONDA", "_PYI", "PYINSTALLER"))
            and key.upper() not in ("VIRTUAL_ENV", "PATH", "TEMP", "TMP")
        }
        env.update(PATH=str(system32), TEMP=str(extraction), TMP=str(extraction))
        report["isolation"] = {
            "chinese_and_space_paths": True,
            "cwd_differs_from_exe_directory": cwd != portable,
            "python_tcl_tk_conda_environment_removed": True,
            "path_only_system32": True,
            "hidden_launch": True,
        }
        report["stage"] = "exe-self-test"
        self_test_log = output / "portable-self-test.txt"
        with (output / "portable-process.txt").open("x", encoding="utf-8") as process_log:
            process = subprocess.Popen(
                [str(exe), "--self-test", "--self-test-log", str(self_test_log)],
                cwd=cwd, env=env, stdout=process_log, stderr=subprocess.STDOUT,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            try:
                report["self_test_exit_code"] = process.wait(timeout=TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                report["timed_out"] = True
                # Only the process launched above and its descendants are targeted.
                try:
                    subprocess.run(
                        [str(system32 / "taskkill.exe"), "/PID", str(process.pid), "/T", "/F"],
                        stdout=process_log, stderr=subprocess.STDOUT,
                        creationflags=subprocess.CREATE_NO_WINDOW, timeout=10,
                    )
                except (OSError, subprocess.TimeoutExpired) as cleanup_error:
                    report["timeout_cleanup_error_type"] = type(cleanup_error).__name__
                finally:
                    if process.poll() is None:
                        process.kill()
                    process.wait(timeout=10)
                raise
        require(report["self_test_exit_code"] == 0, "EXE self-test returned nonzero")
        require("v1.1.1 self-test passed" in self_test_log.read_text(encoding="utf-8-sig"),
                "Expected self-test success text is missing")
        report["self_test_log_sha256"] = sha(self_test_log)
        report["zip_unchanged"] = sha(archive_path) == expected_zip
        report["extracted_exe_unchanged"] = sha(exe) == expected_exe
        require(report["zip_unchanged"] and report["extracted_exe_unchanged"],
                "ZIP or extracted EXE changed during acceptance")
        report.update(passed=True, stage="complete")
    except Exception as error:
        # Do not put environment contents or arbitrary local filenames in public JSON.
        report["error_type"] = type(error).__name__
    finally:
        report["elapsed_seconds"] = round(time.monotonic() - started, 3)
        (output / "release-download-verification.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
    print(json.dumps(report, ensure_ascii=True))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
