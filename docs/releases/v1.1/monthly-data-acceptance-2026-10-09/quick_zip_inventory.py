"""Read-only package inventory without expanding worksheet cell XML."""
import argparse
import hashlib
import json
import posixpath
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


def digest(path):
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('source', type=Path)
    parser.add_argument('--report', required=True, type=Path)
    args = parser.parse_args()
    started = time.perf_counter()
    before = digest(args.source)
    report = {'recorded_utc': datetime.now(timezone.utc).isoformat(),
              'source_name': args.source.name, 'source_bytes': args.source.stat().st_size,
              'source_sha256_before': before, 'script_sha256': digest(Path(__file__)),
              'method': 'ZIP central directory and small workbook/relationship metadata only; worksheet XML not expanded.',
              'privacy': 'No cell values, sheet names, formulas, external URLs or custom formats recorded.'}
    with ZipFile(args.source) as archive:
        report['parts'] = [{'part': i.filename, 'compressed_bytes': i.compress_size,
                            'uncompressed_bytes': i.file_size, 'crc32': format(i.CRC, '08x'),
                            'compression': i.compress_type} for i in archive.infolist()]
        rels = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        targets = {r.get('Id'): (r.get('Target').lstrip('/') if r.get('Target').startswith('/') else
                   posixpath.normpath(posixpath.join('xl', r.get('Target'))))
                   for r in rels if r.get('Type', '').endswith('/worksheet')}
        root = ET.fromstring(archive.read('xl/workbook.xml'))
        report['sheets'] = []
        for node in root.iter():
            if node.tag.rsplit('}', 1)[-1] != 'sheet':
                continue
            rid = next(v for k, v in node.attrib.items() if k.endswith('}id'))
            part = targets[rid]
            info = archive.getinfo(part)
            report['sheets'].append({'index': len(report['sheets']) + 1, 'part': part,
                                     'state': node.get('state', 'visible'),
                                     'sheet_name_sha256': hashlib.sha256(node.get('name', '').encode()).hexdigest(),
                                     'xml_bytes': info.file_size, 'compressed_bytes': info.compress_size})
        patterns = {'media': r'^xl/media/', 'drawings': r'^xl/drawings/[^/]+\.xml$',
                    'charts': r'^xl/charts/[^/]+\.xml$', 'external_links': r'^xl/externalLinks/',
                    'macros': r'vbaProject\.bin$', 'metadata': r'metadata\.xml$',
                    'rich_data': r'^xl/richData/', 'tables': r'^xl/tables/',
                    'comments': r'^xl/comments', 'document_signatures': r'^_xmlsignatures/'}
        report['package_features'] = {k: sum(bool(re.search(p, n)) for n in archive.namelist())
                                      for k, p in patterns.items()}
    report['source_sha256_after'] = digest(args.source)
    report['source_unchanged'] = before == report['source_sha256_after']
    report['seconds'] = time.perf_counter() - started
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: report[k] for k in ('source_bytes', 'source_sha256_before', 'sheets',
                                           'package_features', 'source_unchanged', 'seconds')}, ensure_ascii=True))


if __name__ == '__main__':
    main()
