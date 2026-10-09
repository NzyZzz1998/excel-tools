"""Copy only verified EXE outputs into a clean local delivery directory."""
import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path

EVIDENCE = Path(__file__).resolve().parent
REPO = EVIDENCE.parents[3]


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(name):
    return json.loads((EVIDENCE / name).read_text(encoding='utf-8-sig'))


def main():
    exe = read('exe-content-validation-fixed.json')
    office = read('excel-client-exe-fixed.json')
    engine = read('engine-content-validation-fixed.json')
    inventory = read('input-inventory.json')
    manifest = read('fixed-build/verification-manifest.json')
    assert exe['passed'] and office['passed'] and engine['passed']
    package = REPO / 'dist/ExcelTools-v1.1-Windows-x64.zip'
    package_sha = sha(package)
    assert package_sha == manifest['package']['Hash'].lower()
    assert package_sha == exe['candidate_identity']['dist/ExcelTools-v1.1-Windows-x64.zip']
    assert package_sha == office['candidate_package_sha256']
    source_checks = []
    snapshot_root = Path(manifest['source_snapshot'])
    for record in manifest['source_hashes']:
        snapshot = Path(record['Path'])
        if snapshot.suffix != '.py':
            continue
        live = REPO / snapshot.relative_to(snapshot_root)
        current = sha(live)
        assert current == record['Hash'].lower()
        source_checks.append({'path': str(live), 'sha256': current, 'matches_frozen_build': True})
    original_checks = []
    for record in inventory['files']:
        current = sha(record['source_path'])
        assert current == record['source_sha256_before'].lower()
        original_checks.append({'path': record['source_path'], 'sha256': current, 'unchanged': True})
    engine_outputs = {(r['file'], r['mode']): r['output_sha256'] for r in engine['results']}
    delivery = REPO / 'testfile/验收_v1.1_2026-10-09/交付结果'
    if delivery.exists():
        raise RuntimeError('Preserve existing delivery directory: ' + str(delivery))
    modes = {'default': '默认保留横向合并', 'all': '全部拆分'}
    results = []
    for record in exe['results']:
        source = Path(record['output'])
        actual = sha(source)
        assert actual == record['output_sha256']
        assert actual == engine_outputs[(record['file'], record['mode'])]
        dest = delivery / modes[record['mode']] / source.name
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, dest)
        assert sha(dest) == actual
        results.append({'file': record['file'], 'mode': record['mode'], 'exe_output': str(source),
                        'delivered_output': str(dest), 'sha256': actual,
                        'copied_bytes_match': True, 'matches_source_engine_output': True})
    (delivery / '结果说明.txt').write_text(
        'v1.1 业务验收结果（2026-10-09）\n\n'
        '默认保留横向合并：5 份结果，拆分 79 个区域，填充 1954 格，保留 5 个横向汇总合并。\n'
        '全部拆分：5 份结果，拆分 84 个区域，填充 1972 格，剩余合并为 0。\n'
        '两组均由修复后的实际 ExcelTools v1.1 EXE 生成，已逐格核对并经本机 Excel 读取验证。\n'
        '默认组适合保留原报表横向结构；全部拆分组适合需要消除全部合并的后续处理。\n'
        '原业务文件未修改。源报表原有数字和汇总值保持，不重新计算业务指标。\n'
        '完整报告：E:/codex/excel-tools/docs/releases/v1.1/business-acceptance-2026-10-09/report.md\n',
        encoding='utf-8-sig')
    report = {'recorded_at': datetime.now(timezone.utc).isoformat(), 'passed': True,
              'package': str(package), 'package_sha256': package_sha,
              'source_build_checks': source_checks, 'original_checks': original_checks,
              'outputs': results, 'delivery_dir': str(delivery),
              'office_owned_process_exited': office['owned_excel_exited']}
    evidence = EVIDENCE / 'delivery-integrity.json'
    if evidence.exists():
        raise RuntimeError('Do not overwrite evidence')
    evidence.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'passed': True, 'delivered_files': len(results),
                      'delivery_dir': str(delivery), 'originals_unchanged': len(original_checks)}, ensure_ascii=True))


if __name__ == '__main__':
    main()
