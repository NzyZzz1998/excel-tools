"""Scan current tracked/untracked candidates, plus intentional text build evidence.

Read-only for repository inputs. Does not include ignored testfile/dist data and
never emits candidate secret values. Historical scan remains a separate snapshot.
"""
import argparse
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path

import audit_public_history as history


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    assert not args.report.exists(), 'Preserve earlier evidence'
    repo = history.REPO
    paths = {p.decode('utf-8') for p in history.git('ls-files', '-z', '--cached', '--others', '--exclude-standard').split(b'\0') if p}
    # Parent may force-add otherwise ignored text build logs; include these now.
    for folder in ('docs/releases/v1.1.1', 'docs/reviews/v1.1-post-release-2026-10-09'):
        paths.update(p.relative_to(repo).as_posix() for p in (repo / folder).rglob('*')
                     if p.is_file() and '__pycache__' not in p.parts and p.suffix in ('.txt', '.log', '.json', '.jsonl'))
    secrets, values, images, files, suspicious = [], [], [], [], []
    forbidden = {'.xlsx', '.xlsm', '.xls', '.csv', '.tsv', '.zip', '.7z', '.exe', '.pfx', '.pem', '.key', '.db', '.sqlite'}
    for name in sorted(paths):
        path = repo / name
        if not path.is_file():
            continue
        raw = path.read_bytes()
        location = {'current_path': name}
        secrets.extend(history.scan_text(raw, location))
        values.extend(history.json_candidates(raw, location))
        item = {'path': name, 'bytes': len(raw), 'sha256': hashlib.sha256(raw).hexdigest()}
        files.append(item)
        if path.suffix.lower() in forbidden or path.name.startswith('.env') or name.startswith(('testfile/', 'dist/')):
            suspicious.append(name)
        if history.decode(raw) is None:
            images.append({**item, 'png_signature': raw.startswith(b'\x89PNG\r\n\x1a\n')})
    report = {'recorded_utc': datetime.now(timezone.utc).isoformat(),
              'head': history.git('rev-parse', 'HEAD').decode().strip(),
              'scope': 'Current tracked and nonignored untracked files plus v1.1.1/post-release text logs; no remote calls.',
              'files': files, 'count': len(files), 'credential_candidates': secrets,
              'raw_json_value_candidates': values, 'binary_files': images,
              'business_archive_key_or_ignored_tree_paths': suspicious,
              'limits': ['Concurrent future additions/changes are not covered; final index should be compared to this manifest.',
                         'Pattern scan requires human provenance checks for JSON values and visual image review.']}
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: report[k] for k in ('count', 'credential_candidates', 'raw_json_value_candidates', 'business_archive_key_or_ignored_tree_paths')}, ensure_ascii=True))


if __name__ == '__main__':
    main()
