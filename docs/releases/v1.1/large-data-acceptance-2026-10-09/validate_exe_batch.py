"""Run the independent oracle on a completed six-file EXE acceptance batch."""
import argparse
import contextlib
import io
import json
from datetime import datetime, timezone
from pathlib import Path

from inventory_large import sha_file
from validate_large import run

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--mode',choices=('default','all'),required=True)
    args=parser.parse_args()
    evidence=Path(__file__).resolve().parent
    repo=evidence.parents[3]
    source_dir=repo/'testfile/期间缺货Top20'
    outputs=repo/'testfile/大文件验收_v1.1_2026-10-09/exe-check'/args.mode
    report_dir=evidence/'exe-content'/args.mode
    report_dir.mkdir(parents=True,exist_ok=False)
    manifest_path=evidence/'build/attempt-2/verification-manifest.json'
    manifest=json.loads(manifest_path.read_text(encoding='utf-8-sig'))
    names=('供应商缺货Top20 (1).xlsx','库存满足率 (1).xlsx','期间缺货Top20.xlsx',
           '每日缺货Top50 (1).xlsx','缺货数据统计.xlsx','日明细表 (2).xlsx')
    package_path=Path(manifest['package']['Path'])
    assert sha_file(package_path)==manifest['package']['Hash'].lower(), 'Candidate package identity changed.'
    summary={'recorded_utc':datetime.now(timezone.utc).isoformat(),'mode':args.mode,
             'package_sha256':manifest['package']['Hash'].lower(),
             'executable_sha256':manifest['executable']['Hash'].lower(),
             'build_manifest_sha256':sha_file(manifest_path),
             'validator_sha256':sha_file(evidence/'validate_large.py'),
             'source_scope':'Six authorized original inputs; completed GUI outputs only.',
             'results':[]}
    for index,name in enumerate(names,1):
        original=source_dir/name
        copy=outputs/name
        output=outputs/(Path(name).stem+'_拆分填充.xlsx')
        assert original.is_file() and copy.is_file() and output.is_file(), 'Expected completed batch file missing.'
        assert sha_file(original)==sha_file(copy), 'Input copy differs from original.'
        report_path=report_dir/(str(index).zfill(2)+'.json')
        with contextlib.redirect_stdout(io.StringIO()):
            passed=run(original,output,args.mode=='all',report_path)
        result=json.loads(report_path.read_text(encoding='utf-8'))
        summary['results'].append({'file':name,'passed':passed,'report':str(report_path.relative_to(evidence)),
                                   'report_sha256':sha_file(report_path),'output_sha256':result['output_sha256'],
                                   'source_sha256':result['source_sha256'],'input_copy_matches_original':True,
                                   'cells_checked':sum(sheet['output_cells_checked'] for sheet in result.get('sheets',[])),
                                   'issues':result['issue_count'],'seconds':result['seconds']})
        print(json.dumps(summary['results'][-1],ensure_ascii=True),flush=True)
    summary['passed']=all(item['passed'] for item in summary['results'])
    summary['file_count']=len(summary['results'])
    summary['total_cells_checked']=sum(item['cells_checked'] for item in summary['results'])
    (report_dir/'summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'mode':args.mode,'passed':summary['passed'],'files':summary['file_count'],
                      'cells':summary['total_cells_checked']},ensure_ascii=True),flush=True)
    raise SystemExit(0 if summary['passed'] else 1)

if __name__=='__main__':main()
