"""Repeat the real detail worksheet across sheets for resource testing, not business totals."""
import argparse
import copy
import json
import shutil
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from measure_real import REPO, SOURCE, EVIDENCE, sha


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--copies', type=int, default=2)
    args = parser.parse_args()
    if not 2 <= args.copies <= 31:
        raise ValueError('Use 2 to 31 sheets in a bounded local load cohort.')
    folder = REPO / 'testfile/大文件验收_v1.1_2026-10-09' / ('multisheet-input-' + str(args.copies))
    folder.mkdir(parents=True, exist_ok=False)
    output = folder / ('detail-' + str(args.copies) + '-sheets.xlsx')
    before = sha(SOURCE)
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    with ZipFile(SOURCE) as source, ZipFile(output, 'w') as result:
        workbook = ET.fromstring(source.read('xl/workbook.xml'))
        relationships = ET.fromstring(source.read('xl/_rels/workbook.xml.rels'))
        content_types = ET.fromstring(source.read('[Content_Types].xml'))
        sheets = workbook.find('{' + ns + '}sheets')
        extras = []
        for i in range(1, args.copies):
            number = i + 2
            part = 'xl/worksheets/load' + str(number) + '.xml'
            relation = 'loadSheet' + str(number)
            ET.SubElement(sheets, '{' + ns + '}sheet', {'name': 'Load simulation ' + str(i + 1), 'sheetId': str(number), '{' + rel + '}id': relation})
            ET.SubElement(relationships, relationships[0].tag, {'Id': relation, 'Type': rel + '/worksheet', 'Target': 'worksheets/load' + str(number) + '.xml'})
            template = next(n for n in content_types if n.get('PartName') == '/xl/worksheets/sheet1.xml')
            ET.SubElement(content_types, template.tag, {'PartName': '/' + part, 'ContentType': template.get('ContentType')})
            extras.append(part)
        updates = {'xl/workbook.xml': ET.tostring(workbook, encoding='utf-8', xml_declaration=True),
                   'xl/_rels/workbook.xml.rels': ET.tostring(relationships, encoding='utf-8', xml_declaration=True),
                   '[Content_Types].xml': ET.tostring(content_types, encoding='utf-8', xml_declaration=True)}
        result.comment = source.comment
        for info in source.infolist():
            if info.filename in updates:
                result.writestr(copy.copy(info), updates[info.filename])
            else:
                with source.open(info) as incoming, result.open(copy.copy(info), 'w') as outgoing:
                    shutil.copyfileobj(incoming, outgoing, 1024 * 1024)
        original = source.getinfo('xl/worksheets/sheet1.xml')
        for part in extras:
            info = copy.copy(original)
            info.filename = part
            with source.open(original) as incoming, result.open(info, 'w') as outgoing:
                shutil.copyfileobj(incoming, outgoing, 1024 * 1024)
    record = {'source_sha256': before, 'source_unchanged': sha(SOURCE) == before,
              'copies': args.copies, 'detail_rows_total': 190566 * args.copies,
              'output': str(output), 'output_sha256': sha(output), 'bytes': output.stat().st_size,
              'scope': 'Same daily detail repeated on separate sheets to exercise resources and cross-sheet output; not real monthly transactions or valid monthly totals.'}
    assert record['source_unchanged']
    (EVIDENCE / ('multisheet-input-' + str(args.copies) + '.json')).write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(record, ensure_ascii=True))


if __name__ == '__main__':
    main()
