"""Independent synthetic DOM/stream differential and package failure review."""
import argparse
import copy
import hashlib
import importlib.util
import io
import json
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from validate_large import canonical

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def sheet(rows, merges, attrs='', extra=''):
    return ('<?xml version="1.0" encoding="utf-8" standalone="yes"?>'
            '<worksheet xmlns="'+NS+'"'+attrs+'><dimension ref="A1"/><sheetData>'+rows+
            '</sheetData><mergeCells>'+''.join('<mergeCell ref="'+ref+'"/>' for ref in merges)+
            '</mergeCells>'+extra+'</worksheet>').encode()

def cell(ref, value='anchor', attrs=''):
    return '<c r="'+ref+'" t="inlineStr"'+attrs+'><is><t>'+value+'</t></is></c>'

def hash_file(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def package(path, sheets):
    book='<workbook xmlns="'+NS+'" xmlns:r="'+REL+'"><sheets>'
    rels='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    parts={}
    for index,raw in enumerate(sheets,1):
        book+='<sheet name="Synthetic'+str(index)+'" sheetId="'+str(index)+'" r:id="r'+str(index)+'"/>'
        rels+='<Relationship Id="r'+str(index)+'" Type="'+REL+'/worksheet" Target="worksheets/sheet'+str(index)+'.xml"/>'
        parts['xl/worksheets/sheet'+str(index)+'.xml']=raw
    parts['xl/workbook.xml']=(book+'</sheets></workbook>').encode()
    parts['xl/_rels/workbook.xml.rels']=(rels+'</Relationships>').encode()
    parts['sentinel.bin']=bytes(range(256))
    with ZipFile(path,'w') as archive:
        archive.comment=b'archive-comment'
        for name,raw in parts.items():
            info=ZipInfo(name,(2021,2,3,4,5,6))
            info.compress_type=ZIP_DEFLATED
            info.comment=b'member-comment'
            info.extra=b'\xfe\xca\x02\x00xy'
            info.external_attr=0o100640 << 16
            info.create_system=3
            archive.writestr(info,raw)
    return path

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--engine',type=Path,required=True)
    parser.add_argument('--report',type=Path,required=True)
    args=parser.parse_args()
    with tempfile.TemporaryDirectory(prefix='excel-stream-independent-') as directory:
        folder=Path(directory)
        snapshot=folder/'engine_snapshot.py'
        shutil.copy2(args.engine,snapshot)
        spec=importlib.util.spec_from_file_location('candidate_engine',snapshot)
        engine=importlib.util.module_from_spec(spec);spec.loader.exec_module(engine)
        report={'captured_utc':datetime.now(timezone.utc).isoformat(),'engine_sha256':hash_file(snapshot),
                'script_sha256':hash_file(Path(__file__)),'cases':[],'package_checks':[]}
        cases={
            'sparse_rectangle':sheet('<row r="1" spans="1:4">'+cell('A1')+'<c r="D1"><v>0</v></c></row><row r="4">'+cell('C4','untouched')+'</row>',['A1:C3']),
            'unordered_rows_and_cells':sheet('<row r="3">'+cell('C3')+'</row><row r="1">'+cell('D1')+cell('A1')+'</row>',['A1:A2']),
            'overlap_with_blank_anchor':sheet('<row r="1">'+cell('A1')+'</row>',['A1:A3','A2:A4']),
            'duplicate_merge':sheet('<row r="1">'+cell('A1')+'</row>',['A1:A3','A1:A3']),
            'disjoint_nested_row_spans':sheet('<row r="1">'+cell('A1')+cell('C1')+'</row><row r="2">'+cell('D2')+'</row>',['A1:A4','C1:C2','D2:D3']),
            'character_fidelity':sheet('<row r="1">'+cell('A1','first&#13;second&#10;third&#9;last')+'</row><row r="2">'+cell('B2','outside&#13;value')+'</row>',['A1:A2'],extra='<dataValidations><dataValidation prompt="a&#10;b&#13;c&#9;d"/></dataValidations>'),
            'inherited_space':sheet('<row r="1" xml:space="preserve">'+cell('A1','  padded  ')+'</row><row r="2" xml:space="default">'+cell('B2',' untouched ')+'</row>',['A1:A2']),
            'prefixed_shadowing':('<s:worksheet xmlns:s="'+NS+'"><s:sheetData><s:row r="1"><s:c r="A1" t="inlineStr"><s:is><s:t>anchor</s:t></s:is></s:c></s:row><x:row xmlns:x="'+NS+'" xmlns:s="urn:other" r="2"><x:c r="B2"/></x:row></s:sheetData><s:mergeCells><s:mergeCell ref="A1:A2"/></s:mergeCells></s:worksheet>').encode(),
            'local_payload_namespace':sheet('<row r="1"><c r="A1" xmlns:s="'+NS+'" t="inlineStr"><s:is><s:t>anchor</s:t></s:is></c></row>',['A1:A2']),
            'formula_rejection':sheet('<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>',['A1:A2']),
            'metadata_rejection':sheet('<row r="1">'+cell('A1',attrs=' vm="0"')+'</row>',['A1:A2']),
            'hidden_target_rejection':sheet('<row r="1">'+cell('A1')+'</row><row r="2">'+cell('A2','hidden')+'</row>',['A1:A2']),
            'empty_anchor':sheet('<row r="3"><c r="D3"><v>0</v></c></row>',['A1:B2']),
            'no_match_omitted_coordinates':sheet('<row><c><v>0</v></c></row>',[]),
            'horizontal_only_omitted_coordinates':sheet('<row><c><v>0</v></c></row>',['A1:B1']),
            'cdata_pi_comment':('<?xml version="1.0"?><?probe keep?><worksheet xmlns="'+NS+'"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t><![CDATA[  x<&\r\ny  ]]></t></is></c><!--row-comment--></row></sheetData><mergeCells><mergeCell ref="A1:A2"/></mergeCells><!--tail-comment--></worksheet>').encode(),
        }
        for name,raw in cases.items():
            for all_merges in (False,True):
                outputs={}
                for method in ('dom','stream'):
                    try:
                        if method=='dom':
                            data,count,filled=engine.transform_sheet(raw,all_merges)
                        else:
                            stream=io.BytesIO()
                            count,filled=engine.transform_sheet_stream(io.BytesIO(raw),stream,all_merges)
                            data=stream.getvalue() if count else raw
                        parsed=canonical(ET.fromstring(data))
                        outputs[method]={'status':'ok','stats':[count,filled],'canonical':parsed,
                                         'comment_pi_tokens':{token:token.encode() in data for token in ('<?probe keep?>','<!--row-comment-->','<!--tail-comment-->')}}
                    except Exception as error:
                        outputs[method]={'status':'error','type':type(error).__name__,'message':str(error)}
                same=outputs['dom']==outputs['stream']
                brief={key:{field:value for field,value in value.items() if field!='canonical'} for key,value in outputs.items()}
                report['cases'].append({'name':name,'all_merges':all_merges,'equivalent':same,'outcomes':brief})
        # Force both small worksheet parts through the public streaming branch.
        real_temp=engine.TemporaryFile
        opened=[]
        def tracked_temp(*a,**kw):
            stream=real_temp(*a,**kw);opened.append(stream);return stream
        valid=cases['character_fidelity']
        with patch.object(engine,'STREAM_THRESHOLD',0),patch.object(engine,'TemporaryFile',side_effect=tracked_temp):
            source=package(folder/'metadata.xlsx',[valid,cases['no_match_omitted_coordinates']])
            before_hash=hash_file(source)
            output,stats=engine.process_file(source)
            metadata_equal=True
            with ZipFile(source) as before,ZipFile(output) as after:
                for left,right in zip(before.infolist(),after.infolist()):
                    fields=('filename','date_time','compress_type','comment','extra','external_attr','internal_attr','create_system')
                    metadata_equal &= all(getattr(left,key)==getattr(right,key) for key in fields)
                archive_metadata=before.namelist()==after.namelist() and before.comment==after.comment
                untouched=all(before.read(name)==after.read(name) for name in before.namelist() if name!='xl/worksheets/sheet1.xml')
            report['package_checks'].append({'name':'metadata_untouched_and_no_match','passed':metadata_equal and archive_metadata and untouched and hash_file(source)==before_hash})
            source=package(folder/'reject-later-sheet.xlsx',[valid,cases['formula_rejection']])
            before_files={p.name:hash_file(p) for p in folder.glob('*.xlsx')}
            rejected=False
            try:engine.process_file(source)
            except ValueError:rejected=True
            report['package_checks'].append({'name':'later_sheet_rejection_no_partial_output','passed':rejected and before_files=={p.name:hash_file(p) for p in folder.glob('*.xlsx')}})
            source=package(folder/'write-failure.xlsx',[valid])
            previous=folder/'write-failure_拆分填充.xlsx';previous.write_bytes(b'keep old result')
            before_files={p.name:hash_file(p) for p in folder.glob('*.xlsx')}
            real_open=engine.ZipFile.open
            writes=0
            def failing_open(archive,name,mode='r',*a,**kw):
                nonlocal writes
                if mode=='w':
                    writes+=1
                    if writes==2:raise OSError('synthetic write failure')
                return real_open(archive,name,mode,*a,**kw)
            rejected=False
            with patch.object(engine.ZipFile,'open',new=failing_open):
                try:engine.process_file(source)
                except OSError:rejected=True
            report['package_checks'].append({'name':'late_zip_write_failure_cleanup','passed':rejected and before_files=={p.name:hash_file(p) for p in folder.glob('*.xlsx')}})
        report['package_checks'].append({'name':'temporary_streams_closed','passed':all(stream.closed for stream in opened),'stream_count':len(opened)})
        report['all_differential_cases_equivalent']=all(row['equivalent'] for row in report['cases'])
        report['package_checks_passed']=all(row['passed'] for row in report['package_checks'])
        args.report.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps({'engine_sha256':report['engine_sha256'],'differences':[row for row in report['cases'] if not row['equivalent']], 'package_checks':report['package_checks']},ensure_ascii=False))

if __name__=='__main__':main()
