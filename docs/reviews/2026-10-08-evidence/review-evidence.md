# Quality review evidence — 2026-10-08 Asia/Shanghai

This is a summary of tool executions in the review session, not a CI or official Release log. Business files under testfile were not opened. Project sources were not edited. No network fetch, dependency installation, commit, push, or release was performed.

## Identity
- Workspace: E:/codex/excel-tools
- HEAD and local v1.0.0 tag: 470a4e1b3c5346ed8bdfd3d770b4d235eb45f173
- Source copies and built EXE hashes: verification-manifest.json
- Host: Windows 11, Python 3.12.8, installed PyInstaller 6.20.0.
- CI declares PyInstaller 6.22.3; that version and official GitHub Release artifacts were not exercised.

## Existing tests
Initial `python -m unittest discover -s tests -v` failed during module imports because the system Python did not have openpyxl. No business tests executed in that first invocation.

The installed Codex workspace runtime contains openpyxl 3.1.5. Set process-only PYTHONPATH=C:/Users/win/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/Lib/site-packages and PYTHONDONTWRITEBYTECODE=1, then reran the same command with system Python 3.12.8. Actual result: `Ran 19 tests in 0.518s`, `OK`, exit 0.

`python excel_unmerge_gui.py --self-test`: exit 0.

## Additional temporary synthetic probes
- Used existing `tests/test_excel_unmerge.py` helpers to create a synthetic workbook. Patched ZipFile.writestr to raise OSError at the second archive entry during output. Actual exception: `synthetic disk-write failure`. Source SHA256 remained identical and directory contents matched the pre-operation snapshot, so no output residue remained.
- Added `_xmlsignatures/sig1.xml` to another synthetic archive. process_file rejected with ValueError; no output residue.
- Wrote a synthetic corrupt .xlsx containing `not a zip`. process_file rejected with BadZipFile: `File is not a zip file`; no output residue.
- These were one-off probes; they are not committed regression tests. Full scripted stdout was in the tool transcript; this section records its results.

## Local rebuild and isolated EXE check
Copied the two unchanged source files into this temporary directory, then ran:
`python -m PyInstaller --noconfirm --clean --onefile --windowed --name ExcelTools --distpath <review>/dist --workpath <review>/build --specpath <review> <review>/excel_unmerge_gui.py`
Build exit 0. Full output is build.log.

Copied only the EXE into `免安装 验收`, used separate `独立 工作目录`, removed Python/Tcl/Tk/virtualenv/conda environment variables, restricted PATH to Windows system folders, and set TEMP/TMP to `临时 解压目录`. Launched `--self-test` with Start-Process -WindowStyle Hidden. Actual exit 0 recorded in package-self-test.log; binary identity and limits in verification-manifest.json.

This confirms the local source revision can be bundled and self-tested with the installed toolchain on this host. It does not prove official Release identity, CI success with its pinned version, Windows 10 compatibility, clean-user-machine compatibility, or arbitrary real workbook behavior.

## Review observations
- Existing tests cover 15 conversion/storage cases, 2 batch cases, and 2 actual Tk window cases.
- Write-interruption cleanup, signature rejection, and corrupted ZIP handling had no persisted tests before this review; the above probes passed.
- Workflow triggers are only push main and workflow_dispatch (.github/workflows/windows-build.yml:3-6). No automatic pull_request run is configured. Suggestion: use the same checks for pull requests when that workflow is used.
- Workflow produces an artifact (.github/workflows/windows-build.yml:59-71), while docs/开发说明.md:24 describes manual promotion of the same tested artifact to Release. This is a documented manual publishing boundary, not an automatic publishing bug. A small commit/version/SHA record would make manual promotion auditable; remote status is unverified.
- Windowed self-test diagnostics depend on sys.stderr being present (excel_unmerge_gui.py:326-329); otherwise the exception becomes exit 1 without details. CI surfaces only that exit code (.github/workflows/windows-build.yml:52-53). Suggestion: persist self-test errors to a temporary diagnostics file and collect it on failure.
