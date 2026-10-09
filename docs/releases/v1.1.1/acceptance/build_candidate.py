"""Freeze, test and build v1.1.1 without replacing the published v1.1 files."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import uuid
from datetime import datetime, timezone
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
BUILD_PYTHON = Path('C:/Users/win/AppData/Local/Temp/excel-tools-v1.1-acceptance-23bc299521804363abc3c99f5a99ed02/venv/Scripts/python.exe')
EVIDENCE = ROOT / 'build'


def sha(path):
    with Path(path).open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def run(args, log, cwd, env=None, timeout=180):
    with (EVIDENCE / log).open('x', encoding='utf-8') as output:
        result = subprocess.run(list(map(str, args)), cwd=cwd, env=env,
                                stdout=output, stderr=subprocess.STDOUT, timeout=timeout,
                                creationflags=subprocess.CREATE_NO_WINDOW)
    if result.returncode:
        raise RuntimeError(f'{log}: exit {result.returncode}')


def main():
    package = REPO / 'dist/ExcelTools-v1.1.1-Windows-x64.zip'
    if package.exists() or EVIDENCE.exists():
        raise RuntimeError('Refusing to replace existing build evidence or v1.1.1 ZIP')
    EVIDENCE.mkdir(parents=True)
    work = REPO / 'testfile' / ('v1.1.1-build-' + uuid.uuid4().hex[:8])
    snapshot = work / 'source'
    snapshot.mkdir(parents=True)
    for name in ('excel_unmerge_fill.py', 'excel_unmerge_gui.py'):
        shutil.copy2(REPO / name, snapshot / name)
    shutil.copytree(REPO / 'tests', snapshot / 'tests', ignore=shutil.ignore_patterns('__pycache__'))
    shutil.copy2(REPO / 'packaging/使用说明.txt', snapshot / '使用说明.txt')
    hashes = {str(p.relative_to(snapshot)): sha(p) for p in snapshot.rglob('*') if p.is_file()}
    env = dict(os.environ, PYTHONUTF8='1', PYTHONDONTWRITEBYTECODE='1')
    versions = json.loads(subprocess.check_output([str(BUILD_PYTHON), '-c', 'import sys,PyInstaller,openpyxl,json;print(json.dumps(dict(python=sys.version.split()[0],pyinstaller=PyInstaller.__version__,openpyxl=openpyxl.__version__)))'], env=env))
    assert versions == dict(python='3.12.8', pyinstaller='6.22.3', openpyxl='3.1.5'), versions
    assert 'APP_VERSION = "1.1.1"' in (snapshot / 'excel_unmerge_gui.py').read_text(encoding='utf-8-sig')
    print('Frozen source; running full tests', flush=True)
    run([BUILD_PYTHON, '-m', 'unittest', 'discover', '-s', 'tests', '-v'], 'unit-tests.txt', snapshot, env)
    independent = work / 'independent'
    independent.mkdir()
    shutil.copy2(ROOT / 'independent_acceptance.py', independent)
    run([BUILD_PYTHON, independent / 'independent_acceptance.py', snapshot], 'independent-acceptance.txt', snapshot, env)
    run([BUILD_PYTHON, snapshot / 'excel_unmerge_gui.py', '--self-test', '--self-test-log', EVIDENCE / 'source-self-test.txt'], 'source-process.txt', snapshot, env)
    print('Source verified; building portable EXE', flush=True)
    run([BUILD_PYTHON, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onefile', '--windowed', '--name', 'ExcelTools', '--distpath', work / 'built', '--workpath', work / 'build', '--specpath', snapshot, snapshot / 'excel_unmerge_gui.py'], 'pyinstaller.txt', snapshot, env)
    portable = work / '免安装 验收'
    cwd = work / '独立 工作目录'
    extraction = work / '临时 解压目录'
    for folder in (portable, cwd, extraction): folder.mkdir()
    exe = portable / 'ExcelTools.exe'
    shutil.copy2(work / 'built/ExcelTools.exe', exe)
    isolated = {k: v for k, v in os.environ.items() if not k.upper().startswith(('PYTHON', 'TCL', 'TK', 'CONDA')) and k.upper() != 'VIRTUAL_ENV'}
    isolated.update(PATH=os.environ['SystemRoot'] + '\\System32;' + os.environ['SystemRoot'], TEMP=str(extraction), TMP=str(extraction))
    run([exe, '--self-test', '--self-test-log', EVIDENCE / 'portable-self-test.txt'], 'portable-process.txt', cwd, isolated, timeout=90)
    assert 'v1.1.1 self-test passed' in (EVIDENCE / 'portable-self-test.txt').read_text(encoding='utf-8-sig')
    assert hashes == {str(p.relative_to(snapshot)): sha(p) for p in snapshot.rglob('*') if p.is_file() and p.suffix != '.spec'}
    for name in ('excel_unmerge_fill.py', 'excel_unmerge_gui.py'):
        assert sha(REPO / name) == hashes[name], 'Live application changed after freeze'
    instructions = (snapshot / '使用说明.txt').read_text(encoding='utf-8-sig')
    (portable / '使用说明.txt').write_text(instructions, encoding='utf-8-sig')
    with ZipFile(package, 'x', ZIP_DEFLATED) as archive:
        for name in ('ExcelTools.exe', '使用说明.txt'): archive.write(portable / name, name)
    assert sha(REPO / 'dist/ExcelTools-v1.1-Windows-x64.zip') == '07b801517060896c01ea9be2aaf1420784d2e87fe3a5f2ae4ea1b31f5b0488f6'
    manifest = dict(recorded_utc=datetime.now(timezone.utc).isoformat(), app_version='1.1.1', head=subprocess.check_output(['git','rev-parse','HEAD'], cwd=REPO, text=True).strip(), source_status=subprocess.check_output(['git','status','--short'], cwd=REPO, text=True), source_snapshot=str(snapshot), source_hashes=hashes, dependencies=versions, build_python=str(BUILD_PYTHON), executable=dict(path=str(exe), sha256=sha(exe)), package=dict(path=str(package), sha256=sha(package)), portable_self_test_exit=0, source_and_snapshot_equal=True, previous_release_zip_unchanged=True, limits='Local Windows 11; isolated paths/environment, not a clean-machine claim; synthetic tests only at build stage.')
    (EVIDENCE / 'verification-manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(json.dumps(manifest, ensure_ascii=True), flush=True)


if __name__ == '__main__': main()
