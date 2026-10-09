"""Read-only identity/content audit of the previously accepted v1.1 package.

Extracts the two known package members to a fresh temporary directory. Does not
build, execute bundled code, alter the package, or read business workbooks.
"""
import ast
import hashlib
import json
import marshal
import subprocess
import sys
import tempfile
import types
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile

from PyInstaller.archive.readers import CArchiveReader
import pefile

ROOT = Path(__file__).resolve().parents[5]
EVIDENCE = Path(__file__).resolve().parent
FIXED = ROOT / 'docs/releases/v1.1/business-acceptance-2026-10-09/fixed-build'
EXPECTED_HEAD = '96cfbfeab30be280129aa8af9fff74e5f6be2464'
EXPECTED_ZIP = '92870bdc52451341c7268f6e1937a3741ecd3e1012de0cf2e4026c606ccdd4a9'


def sha(data):
    return hashlib.sha256(data).hexdigest()


def read_json(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def code_identity(value):
    if isinstance(value, types.CodeType):
        fields = ('co_argcount', 'co_posonlyargcount', 'co_kwonlyargcount',
                  'co_nlocals', 'co_stacksize', 'co_flags', 'co_code', 'co_consts',
                  'co_names', 'co_varnames', 'co_name', 'co_qualname',
                  'co_firstlineno', 'co_linetable', 'co_exceptiontable',
                  'co_freevars', 'co_cellvars')
        return {name: code_identity(getattr(value, name)) for name in fields}
    if isinstance(value, bytes):
        return {'bytes': value.hex()}
    if isinstance(value, (tuple, list)):
        return [code_identity(item) for item in value]
    if isinstance(value, frozenset):
        return {'frozenset': sorted((code_identity(item) for item in value), key=repr)}
    if value is Ellipsis:
        return {'ellipsis': True}
    if isinstance(value, complex):
        return {'complex': repr(value)}
    return value


manifest = read_json(FIXED / 'verification-manifest.json')
snapshot_hashes = read_json(FIXED / 'source-snapshot-hashes.json')
head = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip()
report = {'captured_utc': datetime.now(timezone.utc).isoformat(), 'head': head,
          'expected_head_matches': head == EXPECTED_HEAD, 'checks': [],
          'python': sys.version.split()[0], 'script_sha256': sha(Path(__file__).read_bytes())}
for item in snapshot_hashes:
    snapshot_path = Path(item['Path'])
    if '__pycache__' in snapshot_path.parts:
        continue
    relative = snapshot_path.relative_to(Path(manifest['source_snapshot']))
    if relative.as_posix() == '使用说明.txt':
        relative = Path('packaging/使用说明.txt')
    current = (ROOT / relative).read_bytes()
    git_blob = subprocess.check_output(['git', 'show', f'{EXPECTED_HEAD}:{relative.as_posix()}'], cwd=ROOT)
    report['checks'].append({
        'path': relative.as_posix(), 'working_sha256': sha(current),
        'snapshot_sha256': item['Hash'].lower(),
        'working_matches_accepted_snapshot': sha(current) == item['Hash'].lower(),
        'git_blob_sha256': sha(git_blob),
        'git_exact_bytes_match_working': git_blob == current,
        'git_lf_normalized_matches_working': git_blob.replace(b'\r\n', b'\n') == current.replace(b'\r\n', b'\n')})

package = ROOT / 'dist/ExcelTools-v1.1-Windows-x64.zip'
package_bytes = package.read_bytes()
temp_root = Path(tempfile.mkdtemp(prefix='excel-tools-v1.1-release-audit-'))
unpack = temp_root / '发布 自检'
unpack.mkdir()
with ZipFile(package) as archive:
    names = archive.namelist()
    unsafe_names = [name for name in names if Path(name).is_absolute() or '..' in Path(name).parts or ':' in name]
    assert not unsafe_names, 'Unsafe package member names.'
    assert sorted(names) == sorted(['ExcelTools.exe', '使用说明.txt']), 'Unexpected package members.'
    bad_crc = archive.testzip()
    assert bad_crc is None, 'ZIP CRC validation failed.'
    exe_bytes = archive.read('ExcelTools.exe')
    instructions = archive.read('使用说明.txt')
    archive.extractall(unpack)
report['package'] = {
    'path': str(package), 'bytes': len(package_bytes), 'sha256': sha(package_bytes),
    'matches_accepted_hash': sha(package_bytes) == EXPECTED_ZIP == manifest['package']['Hash'].lower(),
    'members': names, 'crc_passed': bad_crc is None, 'unsafe_member_names': unsafe_names,
    'only_application_and_instructions': True, 'temporary_extraction': str(unpack)}
report['executable'] = {
    'sha256': sha(exe_bytes),
    'matches_accepted_executable': sha(exe_bytes) == manifest['executable']['Hash'].lower()}
source_instructions = (ROOT / 'packaging/使用说明.txt').read_bytes()
expected_instructions = b'\xef\xbb\xbf' + source_instructions.removeprefix(b'\xef\xbb\xbf')
report['instructions'] = {
    'sha256': sha(instructions), 'utf8_bom': instructions.startswith(b'\xef\xbb\xbf'),
    'exact_source_plus_utf8_bom': instructions == expected_instructions,
    'version_label': instructions.decode('utf-8-sig').splitlines()[0]}

exe = unpack / 'ExcelTools.exe'
pe = pefile.PE(str(exe), fast_load=True)
report['executable']['pe_machine'] = hex(pe.FILE_HEADER.Machine)
report['executable']['pe_subsystem'] = pe.OPTIONAL_HEADER.Subsystem
report['executable']['x64_windows_gui'] = pe.FILE_HEADER.Machine == 0x8664 and pe.OPTIONAL_HEADER.Subsystem == 2
pe.close()
car = CArchiveReader(str(exe))
pyz = car.open_embedded_archive('PYZ.pyz')
binary_code_checks = []
for module in ('excel_unmerge_gui', 'excel_unmerge_fill'):
    packaged_code = marshal.loads(car.extract(module)) if module in car.toc else pyz.extract(module)
    source = subprocess.check_output(['git', 'show', f'{EXPECTED_HEAD}:{module}.py'], cwd=ROOT)
    compiled_code = compile(source, f'{module}.py', 'exec', optimize=0)
    packaged_identity = code_identity(packaged_code)
    source_identity = code_identity(compiled_code)
    binary_code_checks.append({
        'module': module, 'code_equal_except_co_filename': packaged_identity == source_identity,
        'packaged_code_sha256': sha(json.dumps(packaged_identity, sort_keys=True, ensure_ascii=False).encode()),
        'head_compiled_code_sha256': sha(json.dumps(source_identity, sort_keys=True, ensure_ascii=False).encode())})
report['binary_code_checks'] = binary_code_checks
report['embedded_application_data_members'] = [name for name in car.toc
    if name.lower().endswith(('.xlsx', '.xlsm', '.xls', '.csv', '.7z', '.env', '.key', '.pem'))]
gui_ast = ast.parse((ROOT / 'excel_unmerge_gui.py').read_bytes())
versions = [node.value.value for node in gui_ast.body if isinstance(node, ast.Assign)
            and any(isinstance(target, ast.Name) and target.id == 'APP_VERSION' for target in node.targets)]
report['app_versions'] = versions
report['historical_self_test'] = {
    'exit_code': manifest['portable_self_test_exit'],
    'log': (FIXED / 'portable-self-test.txt').read_text(encoding='utf-8-sig').strip(),
    'process_log': (FIXED / 'portable-process.txt').read_text(encoding='utf-8-sig').strip(),
    'build_script_sha256_matches_manifest': sha((FIXED / 'run_fixed_build.ps1').read_bytes()) == manifest['build_script']['Hash'].lower()}
checks = [report['expected_head_matches'], report['package']['matches_accepted_hash'],
          report['executable']['matches_accepted_executable'], report['executable']['x64_windows_gui'],
          report['instructions']['exact_source_plus_utf8_bom'], versions == ['1.1'],
          not report['embedded_application_data_members']]
checks += [item['working_matches_accepted_snapshot'] and item['git_lf_normalized_matches_working'] for item in report['checks']]
checks += [item['code_equal_except_co_filename'] for item in binary_code_checks]
report['status'] = 'passed' if all(checks) else 'failed'
report['boundaries'] = [
    'Only co_filename is excluded from embedded code comparison; bytecode, constants, names and line/exception tables are included.',
    'Raw Git LF blobs and Windows working-tree CRLF are recorded separately.',
    'This does not replace malware scanning or clean Windows/Office compatibility validation.',
    'No CI artifact is substituted for this package; the accepted local binary retains its existing identity.']
(EVIDENCE / 'package-audit.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'status': report['status'], 'source_checks': len(report['checks']),
                  'binary_code_checks': binary_code_checks, 'temporary_extraction': str(unpack)}, ensure_ascii=False))
raise SystemExit(0 if report['status'] == 'passed' else 1)
