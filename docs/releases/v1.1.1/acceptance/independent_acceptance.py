"""Independent v1.1.1 acceptance (permission injection adapted for .part creation); synthetic data only, no user workbooks."""
import hashlib
import importlib.util
import json
import queue
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(sys.argv.pop(1)).resolve()
sys.path.insert(0, str(ROOT))
import excel_unmerge_fill as engine
import excel_unmerge_gui as gui
import tkinter as tk

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def worksheet(rows, merges, extra=""):
    return (f'<worksheet xmlns="{NS}"><dimension ref="A1:C5"/><sheetData>'
            + rows + '</sheetData><mergeCells count="' + str(len(merges)) + '">'
            + ''.join('<mergeCell ref="' + ref + '"/>' for ref in merges)
            + '</mergeCells>' + extra + '</worksheet>').encode('utf-8')


def basic_sheet():
    return worksheet('<row r="1"><c r="A1" t="inlineStr"><is><t>acceptance</t></is></c></row>', ['A1:A3'])


def workbook(path, raw):
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="{PKG}"><Relationship Id="wb" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{NS}" xmlns:r="{REL}"><sheets><sheet name="独立验收" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': raw,
        'opaque/sentinel.bin': bytes(range(256)),
    }
    with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
        archive.comment = b'acceptance-archive-comment'
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


def parse_sheet(path):
    with ZipFile(path) as archive:
        return ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))


class IndependentAcceptance(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='excel-v1.1-independent-')
        self.folder = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_r1_character_fidelity_and_unmodified_archive_parts(self):
        rows = ('<row r="1"><c r="A1" t="inlineStr"><is><t>first&#13;second&#10;third&#9;tab</t></is></c>'
                '<c r="B1" t="inlineStr"><is><t>outside&#13;value</t></is></c></row>'
                '<row r="3"><c r="B3"><f>SUM(1,2)</f><v>3</v></c></row>')
        extra = '<dataValidations count="1"><dataValidation sqref="C1" prompt="line1&#10;line2&#13;line3&#9;tab"/></dataValidations>'
        source = workbook(self.folder/'characters.xlsx', worksheet(rows, ['A1:A3'], extra))
        before_hash = digest(source)
        before = parse_sheet(source)
        output, stats = engine.process_file(source)
        after = parse_sheet(output)
        expected = 'first\rsecond\nthird\ttab'
        for address in ['A1','A2','A3']:
            self.assertEqual(after.find(f'.//{{{NS}}}c[@r="{address}"]/{{{NS}}}is/{{{NS}}}t').text, expected)
        self.assertEqual(after.find(f'.//{{{NS}}}c[@r="B1"]/{{{NS}}}is/{{{NS}}}t').text, 'outside\rvalue')
        self.assertEqual(after.find(f'.//{{{NS}}}dataValidation').attrib['prompt'], 'line1\nline2\rline3\ttab')
        self.assertEqual(ET.tostring(before.find(f'.//{{{NS}}}c[@r="B3"]')), ET.tostring(after.find(f'.//{{{NS}}}c[@r="B3"]')))
        self.assertEqual(digest(source), before_hash)
        with ZipFile(source) as a, ZipFile(output) as b:
            self.assertEqual(a.namelist(), b.namelist())
            self.assertEqual(a.comment, b.comment)
            for name in a.namelist():
                if name != 'xl/worksheets/sheet1.xml':
                    self.assertEqual(a.read(name), b.read(name))

    def test_r2_metadata_anchors_reject_without_output(self):
        for attribute in ['cm', 'vm', 'extLst']:
            with self.subTest(attribute=attribute):
                cell = ('<c r="A1" t="inlineStr"><is><t>rich</t></is><extLst><ext uri="acceptance"/></extLst></c>'
                        if attribute == 'extLst' else f'<c r="A1" {attribute}="0" t="inlineStr"><is><t>rich</t></is></c>')
                raw = worksheet('<row r="1">' + cell + '</row>', ['A1:A3'])
                source = workbook(self.folder/(attribute+'.xlsx'), raw)
                before = set(self.folder.iterdir())
                before_hash = digest(source)
                with self.assertRaisesRegex(ValueError, '元数据|不支持|复杂|富数据'):
                    engine.process_file(source)
                self.assertEqual(set(self.folder.iterdir()), before)
                self.assertEqual(digest(source), before_hash)

    def test_r1_anchor_local_namespace_and_cdata_remain_parseable(self):
        rows = f'<row r="1"><c r="A1" xmlns:local="{NS}" t="inlineStr"><local:is><local:t><![CDATA[first\rsecond & <third>]]></local:t></local:is></c></row>'
        source = workbook(self.folder/'local-prefix.xlsx', worksheet(rows, ['A1:A3']))
        original_value = parse_sheet(source).find(f'.//{{{NS}}}t').text
        output, _ = engine.process_file(source)
        result = parse_sheet(output)
        for address in ['A1', 'A2', 'A3']:
            self.assertEqual(result.find(f'.//{{{NS}}}c[@r="{address}"]/{{{NS}}}is/{{{NS}}}t').text, original_value)

    def test_r2_hidden_target_metadata_is_not_erased(self):
        raw = worksheet('<row r="1"><c r="A1" t="inlineStr"><is><t>source</t></is></c></row><row r="2"><c r="A2" vm="1"/></row>', ['A1:A3'])
        source = workbook(self.folder/'hidden-metadata.xlsx', raw)
        before = set(self.folder.iterdir())
        before_hash = digest(source)
        with self.assertRaises(ValueError):
            engine.process_file(source)
        self.assertEqual(set(self.folder.iterdir()), before)
        self.assertEqual(digest(source), before_hash)

    def test_r4_corrupt_and_encrypted_inputs_offer_guidance(self):
        for name, raw, expected in [
                ('corrupt.xlsx', b'not-a-zip', '损坏|格式|有效'),
                ('encrypted.xlsx', bytes.fromhex('d0cf11e0a1b11ae1') + bytes(504), '加密|密码')]:
            with self.subTest(name=name):
                source = self.folder/name
                source.write_bytes(raw)
                events = queue.Queue()
                gui.process_batch([str(source)], False, events)
                records = list(events.queue)
                errors = [record[4] for record in records if record[0] == 'result']
                self.assertEqual(len(errors), 1)
                self.assertRegex(errors[0], expected)

    def test_storage_write_failure_cleans_partial_output(self):
        source = workbook(self.folder/'write-failure.xlsx', basic_sheet())
        before = set(self.folder.iterdir())
        before_hash = digest(source)
        real_open = engine.ZipFile.open
        calls = 0
        partial_writes = []
        def fail_second(archive, name, mode='r', *args, **kwargs):
            nonlocal calls
            stream = real_open(archive, name, mode, *args, **kwargs)
            if mode == 'w':
                calls += 1
                if calls == 2:
                    real_write = stream.write
                    def write_partial_then_fail(data):
                        partial_writes.append(real_write(data[:5]))
                        raise OSError('independent injected output failure after partial payload')
                    stream.write = write_partial_then_fail
            return stream
        with patch.object(engine.ZipFile, 'open', fail_second):
            with self.assertRaisesRegex(OSError, 'independent injected output failure'):
                engine.process_file(source)
        self.assertEqual(calls, 2, 'The independent failure injection must be reached')
        self.assertGreater(sum(partial_writes), 0, 'The archive must have received partial payload')
        self.assertEqual(set(self.folder.iterdir()), before)
        self.assertEqual(digest(source), before_hash)

    def test_r4_permission_error_keeps_source_and_gives_recovery_action(self):
        source = workbook(self.folder/'permission.xlsx', basic_sheet())
        before = set(self.folder.iterdir())
        before_hash = digest(source)
        calls = []
        def denied_output(*args, **kwargs):
            calls.append(kwargs)
            raise PermissionError(13, 'independent permission denied', str(kwargs['dir']))
        events = queue.Queue()
        with patch.object(engine, 'mkstemp', denied_output):
            gui.process_batch([str(source)], False, events)
        self.assertEqual(len(calls), 1, 'Temporary-result creation failure must be reached once')
        error = next(record[4] for record in list(events.queue) if record[0] == 'result')
        self.assertIn('可写目录', error)
        self.assertIn('PermissionError', error)
        self.assertIn('independent permission denied', error)
        self.assertEqual(set(self.folder.iterdir()), before)
        self.assertEqual(digest(source), before_hash)

    def test_gui_stop_resume_and_retry_do_not_reprocess_successes(self):
        first = workbook(self.folder/'01-first.xlsx', basic_sheet())
        broken = self.folder/'02-broken.xlsx'
        broken.write_bytes(b'corrupt')
        last = workbook(self.folder/'03-last.xlsx', worksheet('<row r="1"><c r="A1" t="inlineStr"><is><t>header</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>rows</t></is></c></row>', ['A1:C1','A2:A3']))
        fourth = workbook(self.folder/'04-final.xlsx', basic_sheet())
        root = tk.Tk()
        root.attributes('-alpha', 0)
        app = gui.Application(root)
        root.update()
        entered = threading.Event()
        release = threading.Event()
        entered_third = threading.Event()
        release_third = threading.Event()
        real_process = gui.process_file
        processed = []
        def pause_first(path, **kwargs):
            processed.append(Path(path).name)
            if Path(path) == first and processed.count(first.name) == 1:
                entered.set()
                if not release.wait(10):
                    raise TimeoutError('independent acceptance did not release first file')
            if Path(path) == last:
                entered_third.set()
                if not release_third.wait(10):
                    raise TimeoutError('independent acceptance did not release third file')
            return real_process(path, **kwargs)
        def pump_until(predicate, seconds=10):
            end = time.monotonic() + seconds
            while not predicate() and time.monotonic() < end:
                root.update()
                time.sleep(.01)
            root.update()
            self.assertTrue(predicate(), 'GUI condition did not become true')
        try:
            with patch.object(gui, 'process_file', pause_first):
                app.files = tuple(str(p) for p in [first, broken, last, fourth])
                app.start()
                pump_until(lambda: entered.is_set() and first.name in app.status.get())
                app.stop_button.invoke()
                release.set()
                pump_until(lambda: not app.running)
                self.assertEqual(tuple(map(Path, app.pending_files)), (broken, last, fourth))
                self.assertEqual(processed, [first.name])
                self.assertNotIn('disabled', app.resume_button.state())
                app.all_merges.set(True)
                app.resume_button.invoke()
                pump_until(lambda: entered_third.is_set() and last.name in app.status.get())
                app.stop_button.invoke()
                release_third.set()
                pump_until(lambda: not app.running)
                self.assertEqual(tuple(map(Path, app.failed_files)), (broken,))
                self.assertEqual(tuple(map(Path, app.pending_files)), (fourth,))
                self.assertIn('BadZipFile', app.results.get('1.0', 'end'))
                self.assertIn('技术详情：', app.results.get('1.0', 'end'))
                last_output = next(self.folder.glob('03-last_拆分填充*.xlsx'))
                self.assertEqual([n.attrib['ref'] for n in parse_sheet(last_output).findall(f'{{{NS}}}mergeCells/{{{NS}}}mergeCell')], ['A1:C1'])
                workbook(broken, basic_sheet())
                app.retry_button.invoke()
                pump_until(lambda: not app.running)
                self.assertFalse(app.failed_files)
                self.assertEqual(processed, [first.name, broken.name, last.name, broken.name])
                self.assertEqual(tuple(map(Path, app.pending_files)), (fourth,))
                app.resume_button.invoke()
                pump_until(lambda: not app.running)
                self.assertEqual(processed, [first.name, broken.name, last.name, broken.name, fourth.name])
                self.assertFalse(app.pending_files)
                self.assertEqual(len(list(self.folder.glob('01-first_拆分填充*.xlsx'))), 1)
                self.assertEqual(len(list(self.folder.glob('03-last_拆分填充*.xlsx'))), 1)
                self.assertIn('disabled', app.retry_button.state())
        finally:
            release.set()
            release_third.set()
            if app.running:
                end = time.monotonic() + 10
                while app.running and time.monotonic() < end:
                    root.update()
                    time.sleep(.01)
            root.destroy()

    def test_self_test_failure_records_traceback_without_stderr(self):
        log = self.folder/'diagnostics.log'
        with patch.object(sys, 'argv', ['ExcelTools', '--self-test', '--self-test-log', str(log)]), patch.object(sys, 'stderr', None), patch.object(gui, 'self_test', side_effect=RuntimeError('independent-diagnostic-probe')):
            self.assertEqual(gui.main(), 1)
        self.assertIn('independent-diagnostic-probe', log.read_text(encoding='utf-8'))
        self.assertIn('Traceback', log.read_text(encoding='utf-8'))


if __name__ == '__main__':
    result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(IndependentAcceptance))
    if not result.wasSuccessful():
        raise SystemExit(1)
