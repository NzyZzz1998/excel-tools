"""Bounded synthetic check; never terminates an existing user process."""
import hashlib
import json
import subprocess
import sys
import time
import uuid
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, is_zipfile

REPO = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(REPO))
import excel_unmerge_fill as engine


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def child(source, marker):
    original_copy = engine.shutil.copyfileobj

    def pause_inside_output(src, dst, length=0):
        if hasattr(dst, "_zipfile"):
            chunk = src.read(64)
            dst.write(chunk)
            dst._zipfile.fp.flush()
            Path(marker).write_text("output member opened and prefix written", encoding="utf-8")
            time.sleep(45)
            raise RuntimeError("parent failed to terminate owned probe within budget")
        return original_copy(src, dst, length)

    engine.shutil.copyfileobj = pause_inside_output
    engine.process_file(source)


def main():
    scratch = REPO / "testfile" / ("review-hard-stop-" + uuid.uuid4().hex[:8])
    scratch.mkdir(parents=True)
    source = scratch / "synthetic.xlsx"
    marker = scratch / "child-marker.txt"
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with ZipFile(source, "w", ZIP_DEFLATED) as book:
        book.writestr("xl/workbook.xml", '<workbook xmlns="' + ns + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Synthetic" sheetId="1" r:id="rId1"/></sheets></workbook>')
        book.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        book.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="' + ns + '"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>synthetic value</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:A3"/></mergeCells></worksheet>')
    before = digest(source)
    prior, _ = engine.process_file(source)
    prior_before = digest(prior)
    proc = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), "--child", str(source), str(marker)], cwd=REPO, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    try:
        deadline = time.monotonic() + 15
        while not marker.exists() and proc.poll() is None and time.monotonic() < deadline:
            time.sleep(0.02)
        if not marker.exists():
            raise RuntimeError("owned child did not reach output checkpoint")
        proc.terminate()
        stdout, stderr = proc.communicate(timeout=5)
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=5)
    interrupted = scratch / "synthetic_拆分填充_2.xlsx"
    retry, retry_stats = engine.process_file(source)
    data = {
        "scope": "synthetic OOXML package, OS hard termination during final ZIP write; not ordinary Stop button or handled exception",
        "engine_sha256": digest(Path(engine.__file__)),
        "scratch_directory": str(scratch),
        "owned_child_exit_code": proc.returncode,
        "child_stderr": stderr.decode("utf-8", errors="replace"),
        "source_unchanged": digest(source) == before,
        "prior_valid_result_unchanged": digest(prior) == prior_before and is_zipfile(prior),
        "interrupted_output": {"name": interrupted.name, "exists": interrupted.exists(), "bytes": interrupted.stat().st_size, "is_valid_zip": is_zipfile(interrupted)},
        "retry_output": {"name": retry.name, "is_valid_zip": is_zipfile(retry), "stats": retry_stats},
        "partial_output_retained_after_retry": interrupted.exists(),
    }
    assert data["source_unchanged"] and data["prior_valid_result_unchanged"]
    assert not data["interrupted_output"]["is_valid_zip"]
    assert data["retry_output"]["is_valid_zip"] and data["partial_output_retained_after_retry"]
    target = Path(__file__).with_name("hard-stop-result.json")
    target.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--child":
        child(sys.argv[2], sys.argv[3])
    else:
        main()
