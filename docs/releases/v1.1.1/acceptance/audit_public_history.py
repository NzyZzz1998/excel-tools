"""Read-only local Git history scan. Reports locations/classes, never secret values."""
import collections
import io
import json
from pathlib import Path, PurePosixPath
import re
import subprocess
from datetime import datetime, timezone

REPO = Path(__file__).resolve().parents[4]


def git(*args):
    return subprocess.check_output(['git', *args], cwd=REPO)


SECRET_PATTERNS = {
    'private_key_armor': rb'-----BEGIN (?:[A-Z0-9]+ )*PRIVATE KEY-----',
    'github_classic_token': rb'\bgh[pousr]_[A-Za-z0-9]{30,}\b',
    'github_fine_grained_token': rb'\bgithub_pat_[A-Za-z0-9_]{50,}\b',
    'aws_access_key_id': rb'\b(?:AKIA|ASIA)[A-Z0-9]{16}\b',
    'openai_style_key': rb'\bsk-(?:(?:proj|svcacct)-)?[A-Za-z0-9_-]{30,}\b',
    'stripe_live_secret': rb'\bsk_live_[A-Za-z0-9]{16,}\b',
    'slack_token': rb'\bxox[baprs]-[A-Za-z0-9-]{20,}\b',
    'url_embedded_credentials': rb'https?://[^\s/:<>\x22]+:[^\s/@<>\x22]+@[^\s/<>\x22]+',
}
ASSIGNMENT = re.compile(r'''(?i)\b(password|passwd|api[_-]?key|client[_-]?secret|access[_-]?token|auth[_-]?token)\s*[:=]\s*['"]([^'"\r\n]{8,})['"]''')


def decode(raw):
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        return raw.decode('utf-16')
    try:
        return raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        return None


def scan_text(raw, location):
    findings = []
    for name, pattern in SECRET_PATTERNS.items():
        for match in re.finditer(pattern, raw):
            findings.append({'location': location, 'class': name,
                             'line': raw[:match.start()].count(b'\n') + 1})
    text = decode(raw)
    if text is not None:
        for match in ASSIGNMENT.finditer(text):
            value = match[2]
            if any(word in value.lower() for word in ('example', 'placeholder', 'your_', 'your-', 'synthetic', '${', '<')):
                continue
            findings.append({'location': location, 'class': 'credential_assignment_candidate',
                             'key': match[1], 'value_length': len(value),
                             'line': text[:match.start()].count('\n') + 1})
    return findings


def json_candidates(raw, location):
    text = decode(raw)
    if text is None:
        return []
    try:
        data = json.loads(text)
    except (ValueError, TypeError):
        return []
    hits, pending = [], [('', data)]
    names = {'value', 'values', 'cell_values', 'source_values', 'output_values', 'original_values',
             'formula', 'formulas', 'cell_text', 'row_data', 'rows_data', 'raw_data'}
    while pending:
        path, value = pending.pop()
        if isinstance(value, dict):
            for key, item in value.items():
                pointer = path + '/' + str(key)
                if str(key).lower() in names and item not in (None, '', [], {}):
                    hits.append({'location': location, 'json_pointer': pointer,
                                 'value_type': type(item).__name__,
                                 'item_count_or_length': len(item) if hasattr(item, '__len__') else None})
                pending.append((pointer, item))
        elif isinstance(value, list):
            pending.extend((path + '/' + str(n), item) for n, item in enumerate(value))
    return hits


def main():
    commits = git('rev-list', '--all').decode().splitlines()
    blobs, paths = {}, set()
    for commit in commits:
        for record in git('ls-tree', '-r', '-z', commit).split(b'\0'):
            if not record:
                continue
            metadata, raw_path = record.split(b'\t', 1)
            _, kind, oid = metadata.decode().split()
            if kind != 'blob':
                continue
            path = raw_path.decode('utf-8')
            paths.add(path)
            blobs.setdefault(oid, set()).add(path)
    batch = subprocess.run(['git', 'cat-file', '--batch'], cwd=REPO,
                           input=('\n'.join(blobs) + '\n').encode(), stdout=subprocess.PIPE, check=True)
    stream = io.BytesIO(batch.stdout)
    secrets, value_candidates, binary = [], [], []
    blob_sizes = {}
    for requested_oid in blobs:
        header = stream.readline().decode().split()
        oid, kind, size = header[0], header[1], int(header[2])
        assert oid == requested_oid and kind == 'blob'
        raw = stream.read(size)
        assert stream.read(1) == b'\n'
        blob_sizes[oid] = size
        location = {'blob': oid, 'paths': sorted(blobs[oid])}
        secrets.extend(scan_text(raw, location))
        value_candidates.extend(json_candidates(raw, location))
        if decode(raw) is None:
            binary.append({'blob': oid, 'paths': sorted(blobs[oid]), 'bytes': size,
                           'png_signature': raw.startswith(b'\x89PNG\r\n\x1a\n')})
    tracked = [p.decode('utf-8') for p in git('ls-files', '-z').split(b'\0') if p]
    current_missing = []
    for path in tracked:
        absolute = REPO / path
        if not absolute.is_file():
            current_missing.append(path)
            continue
        raw = absolute.read_bytes()
        secrets.extend(scan_text(raw, {'current_tracked_path': path}))
        value_candidates.extend(json_candidates(raw, {'current_tracked_path': path}))
    secrets.extend(scan_text(git('log', '--all', '--format=%H%n%B'), {'scope': 'reachable_commit_messages'}))
    sensitive_suffixes = {'.xlsx', '.xlsm', '.xls', '.csv', '.tsv', '.zip', '.7z', '.tar', '.gz',
                          '.sqlite', '.db', '.pem', '.key', '.pfx', '.env'}
    suspicious_paths = [p for p in sorted(paths) if PurePosixPath(p).suffix.lower() in sensitive_suffixes
                        or PurePosixPath(p).name.startswith('.env') or p.startswith(('testfile/', 'dist/'))]
    report = {'recorded_utc': datetime.now(timezone.utc).isoformat(),
              'head': git('rev-parse', 'HEAD').decode().strip(),
              'refs': git('for-each-ref', '--format=%(refname) %(objectname)').decode().splitlines(),
              'reachable_commits': len(commits), 'unique_history_blobs': len(blobs),
              'historical_paths': len(paths), 'current_tracked_files': len(tracked),
              'current_missing_paths': current_missing,
              'total_unique_blob_bytes': sum(blob_sizes.values()),
              'historical_extension_counts': dict(collections.Counter(PurePosixPath(p).suffix.lower() or '(none)' for p in paths)),
              'business_archive_key_or_ignored_tree_paths': suspicious_paths,
              'credential_candidates': secrets, 'json_raw_value_candidates': value_candidates,
              'binary_blobs': binary,
              'limits': ['All local refs reachable history and current tracked files; no network call.',
                         'Deleted reachable files are included; local unreachable objects, remote-only refs, LFS servers and external attachments are not scanned.',
                         'Credential patterns are heuristic; filenames and structure are not a proof of absence of all sensitive content.',
                         'Images require separate visual review; binary blobs are not OCR-scanned.',
                         'Values matching patterns are never copied into this report.']}
    destination = Path(__file__).with_name('public-history-scan.json')
    assert not destination.exists(), 'Preserve earlier evidence'
    destination.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: report[k] for k in ('head', 'reachable_commits', 'unique_history_blobs',
                      'historical_paths', 'current_tracked_files', 'business_archive_key_or_ignored_tree_paths',
                      'credential_candidates', 'json_raw_value_candidates')}, ensure_ascii=True))


if __name__ == '__main__':
    main()
