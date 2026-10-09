"""Synthetic positive and corruption checks for the independent large oracle."""
import contextlib
import io
import json
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from validate_large import run

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
def sheet(rows, merges='', dimension='A1'):
    return ('<worksheet xmlns="'+NS+'"><dimension ref="'+dimension+'"/><sheetData>'+rows+
            '</sheetData>'+('<mergeCells>'+merges+'</mergeCells>' if merges else '')+
            '<pageMargins left="0.7"/></worksheet>').encode()

def cell(ref, text, kind='inlineStr', style=None):
    attrs = ' r="'+ref+'" t="'+kind+'"'+(' s="'+style+'"' if style else '')
    payload = '<is><t>'+text+'</t></is>' if kind == 'inlineStr' else '<v>'+text+'</v>'
    return '<c'+attrs+'>'+payload+'</c>'

alpha=lambda ref:cell(ref,'synthetic anchor',style='1')
source_rows='<row r="1" spans="1:4">'+alpha('A1')+cell('D1','0','n')+'</row><row r="3" spans="1:3">'+cell('C3','keep')+'</row>'
default_rows='<row r="1">'+alpha('A1')+cell('D1','0','n')+'</row><row r="2">'+alpha('A2')+'</row><row r="3" spans="1:3">'+cell('C3','keep')+'</row>'
all_rows='<row r="1">'+alpha('A1')+alpha('B1')+cell('D1','0','n')+'</row><row r="2">'+alpha('A2')+alpha('B2')+'</row><row r="3">'+cell('C3','keep')+'<c r="F3"/></row>'
book=('<workbook xmlns="'+NS+'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="test1" sheetId="1" r:id="r1"/><sheet name="test2" sheetId="2" r:id="r2"/></sheets></workbook>').encode()
rels=b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'
base={'xl/workbook.xml':book,'xl/_rels/workbook.xml.rels':rels,'xl/styles.xml':b'<styles/>',
      'xl/worksheets/sheet2.xml':sheet('<row r="1">'+cell('C1','1','b')+'</row>')}

def write(path, xml, corrupt_style=False):
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for name,data in base.items():
            z.writestr(name,b'<changed/>' if corrupt_style and name=='xl/styles.xml' else data)
        z.writestr('xl/worksheets/sheet1.xml',xml)

results=[]
with tempfile.TemporaryDirectory(prefix='excel-tools-independent-oracle-') as folder:
    folder=Path(folder)
    source=folder/'source.xlsx'
    write(source,sheet(source_rows,'<mergeCell ref="A1:B2"/><mergeCell ref="E3:F3"/>'))
    default_xml=sheet(default_rows,'<mergeCell ref="A1:B1"/><mergeCell ref="A2:B2"/><mergeCell ref="E3:F3"/>','A1:F3')
    all_xml=sheet(all_rows,'','A1:F3')
    cases=[('default',default_xml,False,False,True),('all',all_xml,True,False,True),
           ('wrong_value',default_xml.replace(b'<t>keep</t>',b'<t>wrong</t>'),False,False,False),
           ('wrong_type',default_xml.replace(b'r="D1" t="n"',b'r="D1" t="s"'),False,False,False),
           ('wrong_target_style',default_xml.replace(b'r="A2" t="inlineStr" s="1"',b'r="A2" t="inlineStr" s="2"'),False,False,False),
           ('missing_cell',default_xml.replace(alpha('A2').encode(),b''),False,False,False),
           ('wrong_dimension',default_xml.replace(b'ref="A1:F3"',b'ref="A1"'),False,False,False),
           ('changed_untouched_part',default_xml,False,True,False)]
    for name,xml,all_merges,corrupt_style,expected in cases:
        output=folder/(name+'.xlsx')
        write(output,xml,corrupt_style)
        report=folder/(name+'.json')
        with contextlib.redirect_stdout(io.StringIO()):
            actual=run(source,output,all_merges,report)
        payload=json.loads(report.read_text(encoding='utf-8'))
        results.append({'scenario':name,'expected_pass':expected,'actual_pass':actual,
                        'test_passed':actual==expected,'issues':payload['issues']})
    # A worksheet without dimension is valid; the contract does not require
    # adding it during processing.
    source_without_dimension=folder/'no-dimension-source.xlsx'
    write(source_without_dimension,sheet(source_rows,'<mergeCell ref="A1:B2"/><mergeCell ref="E3:F3"/>').replace(b'<dimension ref="A1"/>',b''))
    output=folder/'no-dimension-output.xlsx'
    write(output,default_xml.replace(b'<dimension ref="A1:F3"/>',b''))
    report=folder/'no-dimension.json'
    with contextlib.redirect_stdout(io.StringIO()):
        actual=run(source_without_dimension,output,False,report)
    payload=json.loads(report.read_text(encoding='utf-8'))
    results.append({'scenario':'absent_dimension_preserved','expected_pass':True,
                    'actual_pass':actual,'test_passed':actual,'issues':payload['issues']})
    hidden_source=folder/'hidden-source.xlsx'
    hidden_rows=source_rows.replace('</row><row r="3"',cell('B1','hidden')+'</row><row r="3"',1)
    write(hidden_source,sheet(hidden_rows,'<mergeCell ref="A1:B2"/><mergeCell ref="E3:F3"/>'))
    output=folder/'hidden-output.xlsx'
    write(output,all_xml)
    report=folder/'hidden.json'
    with contextlib.redirect_stdout(io.StringIO()):
        actual=run(hidden_source,output,True,report)
    payload=json.loads(report.read_text(encoding='utf-8'))
    results.append({'scenario':'hidden_source_content_must_not_be_overwritten','expected_pass':False,
                    'actual_pass':actual,'test_passed':not actual,'issues':payload['issues']})
summary={'passed':all(item['test_passed'] for item in results),'cases':results}
Path(__file__).with_name('validator-smoke.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(summary,ensure_ascii=False))
raise SystemExit(0 if summary['passed'] else 1)
