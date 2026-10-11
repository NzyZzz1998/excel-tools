"""Verify a newly downloaded v1.2.0 ZIP against release-package.json.

Usage: python verify_release_download.py DOWNLOAD.zip --output-dir NEW_DIRECTORY
No download is performed. The report directory must not already exist.
Only the newly extracted executable is launched; existing releases are untouched.
"""
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
from zipfile import ZipFile


APP_VERSION = "1.2.0"
ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
MANIFEST = ROOT / "release-package.json"
MEMBERS = ("ExcelTools.exe", "使用说明.txt")
TIMEOUT_SECONDS = 90


def sha(path):
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def identity(record):
    digest, size = record["sha256"], record["size"]
    require(isinstance(digest, str) and re.fullmatch(r"[0-9a-fA-F]{64}", digest),
            "Manifest hash is invalid")
    require(isinstance(size, int) and not isinstance(size, bool) and size > 0,
            "Manifest file size is invalid")
    return digest.lower(), size


def public_log(text, paths):
    """Keep diagnostics while removing this run's private absolute paths."""
    replacements = {str(path): label for path, label in paths if path}
    for path, label in sorted(replacements.items(), key=lambda item: -len(item[0])):
        for spelling in {path, path.replace("\\", "/"), path.replace("\\", "\\\\")}:
            text = re.sub(re.escape(spelling), lambda match: label, text, flags=re.IGNORECASE)
    # The self-test needs no process IDs in publicly retained diagnostics.
    return re.sub(r"(?i)\b(?:pid|process\s+id)\s*[:=#]?\s*\d+\b", "<process-id>", text)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("zip_path", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {
        "recorded_utc": datetime.now(timezone.utc).isoformat(),
        "app_version": APP_VERSION,
        "passed": False,
        "stage": "manifest",
        "script_sha256": sha(__file__),
        "timeout_seconds": TIMEOUT_SECONDS,
        "limits": "Local Windows self-test in isolated paths and environment; "
                  "not a clean-machine or business-workbook acceptance test.",
    }
    started = time.monotonic()
    raw_logs = []
    private_paths = [(output, "<report-directory>"), (REPO, "<repository>"),
                     (os.environ.get("USERPROFILE"), "<user-profile>"),
                     (os.environ.get("TEMP"), "<system-temp>"),
                     (os.environ.get("TMP"), "<system-temp>")]
    try:
        require(os.name == "nt", "This release acceptance requires Windows")
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8-sig"))
        require(manifest["version"] == APP_VERSION, "Unexpected manifest version")
        expected_zip, expected_zip_size = identity(manifest["package"])
        expected_exe, expected_exe_size = identity(manifest["executable"])
        sources = manifest["source"]
        require(isinstance(sources, dict) and sources, "Missing frozen source hashes")
        require(all(isinstance(name, str) and re.fullmatch(r"[A-Za-z0-9_]+\.py", name)
                    and isinstance(digest, str) and re.fullmatch(r"[0-9a-fA-F]{64}", digest)
                    for name, digest in sources.items()), "Invalid source hash mapping")
        report.update(manifest_sha256=sha(MANIFEST), source_sha256=sources,
                      expected_zip_sha256=expected_zip, expected_zip_size=expected_zip_size,
                      expected_exe_sha256=expected_exe, expected_exe_size=expected_exe_size)

        report["stage"] = "zip-hash-and-size"
        archive_path = args.zip_path.resolve(strict=True)
        private_paths.append((archive_path, "<downloaded-zip>"))
        report["downloaded_zip_sha256"] = sha(archive_path)
        report["downloaded_zip_size"] = archive_path.stat().st_size
        require(report["downloaded_zip_sha256"] == expected_zip
                and report["downloaded_zip_size"] == expected_zip_size,
                "Downloaded ZIP does not match frozen release")

        report["stage"] = "zip-members-and-crc"
        with ZipFile(archive_path) as archive:
            require(sorted(archive.namelist()) == sorted(MEMBERS),
                    "ZIP must contain exactly ExcelTools.exe and 使用说明.txt")
            require(archive.testzip() is None, "ZIP CRC validation failed")
            require(archive.getinfo("ExcelTools.exe").file_size == expected_exe_size,
                    "ZIP executable size does not match frozen release")
            report["members"] = [
                {"name": item.filename, "size_bytes": item.file_size,
                 "compressed_bytes": item.compress_size, "crc32": f"{item.CRC:08x}"}
                for item in archive.infolist()
            ]
            report["zip_crc_passed"] = True
            test_root = REPO / "testfile"
            test_root.mkdir(exist_ok=True)
            work = Path(tempfile.mkdtemp(prefix=f"v{APP_VERSION}-release-download-", dir=test_root))
            portable = work / "下载 解压验收"
            cwd = work / "独立 工作目录"
            extraction = work / "临时 解压目录"
            for folder in (portable, cwd, extraction):
                folder.mkdir()
            for name in MEMBERS:
                with archive.open(name) as source, (portable / name).open("xb") as target:
                    shutil.copyfileobj(source, target, length=1024 * 1024)
        report["work_directory_relative_to_repo"] = work.relative_to(REPO).as_posix()
        private_paths.extend([(work, "<isolated-work-directory>"),
                              (portable, "<portable-directory>"),
                              (cwd, "<independent-cwd>"),
                              (extraction, "<extraction-temp>")])
        report["stage"] = "exe-hash-and-size"
        exe = portable / "ExcelTools.exe"
        report["extracted_exe_sha256"] = sha(exe)
        report["extracted_exe_size"] = exe.stat().st_size
        require(report["extracted_exe_sha256"] == expected_exe
                and report["extracted_exe_size"] == expected_exe_size,
                "Extracted EXE does not match frozen release")

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
        self_test_log = work / "raw-self-test.txt"
        process_log_path = work / "raw-process.txt"
        raw_logs = [(self_test_log, output / "portable-self-test.txt"),
                    (process_log_path, output / "portable-process.txt")]
        with process_log_path.open("x", encoding="utf-8") as process_log:
            process = subprocess.Popen(
                [str(exe), "--self-test", "--self-test-log", str(self_test_log)],
                cwd=cwd, env=env, stdout=process_log, stderr=subprocess.STDOUT,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            try:
                report["self_test_exit_code"] = process.wait(timeout=TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                report["timed_out"] = True
                # Never enumerate or target executables by name. The only PID
                # passed to taskkill is from our retained Popen child object.
                try:
                    if process.poll() is None:
                        cleanup = subprocess.run(
                            [str(system32 / "taskkill.exe"), "/PID", str(process.pid), "/T", "/F"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                            creationflags=subprocess.CREATE_NO_WINDOW, timeout=10,
                        )
                        report["timeout_tree_cleanup_exit_code"] = cleanup.returncode
                except (OSError, subprocess.TimeoutExpired) as cleanup_error:
                    report["timeout_cleanup_error_type"] = type(cleanup_error).__name__
                finally:
                    if process.poll() is None:
                        process.kill()  # Popen uses the retained Windows process handle.
                    process.wait(timeout=10)
                raise
        require(report["self_test_exit_code"] == 0, "EXE self-test returned nonzero")
        require(f"v{APP_VERSION} self-test passed" in self_test_log.read_text(encoding="utf-8-sig"),
                "Expected versioned self-test success text is missing")
        report["zip_unchanged"] = (sha(archive_path) == expected_zip
                                   and archive_path.stat().st_size == expected_zip_size)
        report["extracted_exe_unchanged"] = (sha(exe) == expected_exe
                                             and exe.stat().st_size == expected_exe_size)
        require(report["zip_unchanged"] and report["extracted_exe_unchanged"],
                "ZIP or extracted EXE changed during acceptance")
        report.update(passed=True, stage="complete")
    except Exception as error:
        # Error text can contain private absolute paths; expose type and stage.
        report["error_type"] = type(error).__name__
    finally:
        for raw, destination in raw_logs:
            try:
                text = raw.read_text(encoding="utf-8-sig", errors="replace") if raw.exists() else ""
                destination.write_text(public_log(text, private_paths), encoding="utf-8")
                report[destination.stem.replace("-", "_") + "_sha256"] = sha(destination)
            except OSError as log_error:
                report.update(passed=False, log_error_type=type(log_error).__name__)
        report["elapsed_seconds"] = round(time.monotonic() - started, 3)
        (output / "release-download-verification.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
    print(json.dumps(report, ensure_ascii=True))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
