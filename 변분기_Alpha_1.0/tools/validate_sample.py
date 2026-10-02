"""Local document check. No Word process unless --word is explicitly supplied."""
import argparse
import json
import logging
from pathlib import Path
import sys
import threading
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from pair_analysis import read_snapshot, compare_snapshots, tracked_advice, digest


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('before', type=Path)
    parser.add_argument('after', type=Path)
    parser.add_argument('--tracked', type=Path)
    parser.add_argument('--word', action='store_true')
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    if args.word and not args.output:
        parser.error('--word requires --output (use an ignored outputs folder)')
    started = time.monotonic()
    before, after = read_snapshot(args.before), read_snapshot(args.after)
    pairs, warnings = compare_snapshots(before, after)
    result = {'before_blocks': len(before.blocks), 'after_blocks': len(after.blocks),
              'changed_blocks': len(pairs), 'equal_text_changed_blocks': sum(
                  p['before'] is not None and p['after'] is not None and p['before'].text == p['after'].text for p in pairs),
              'comparison_seconds': round(time.monotonic()-started, 3), 'warnings': warnings}
    if args.tracked:
        advice = tracked_advice(args.tracked, before, after)
        result['tracked_status'] = advice['status']
        result['tracked_revision_count'] = advice['revision_count']
    if args.word:
        from analyzer import create_analysis_package
        args.output.parent.mkdir(parents=True, exist_ok=True)
        logging.basicConfig(filename=args.output.with_suffix('.log'), encoding='utf-8', level=logging.INFO)
        result['word_integration'] = create_analysis_package(args.before, args.after, args.tracked, args.output,
            lambda e: print(e.get('stage', ''), e.get('current', ''), flush=True), threading.Event(), logging.getLogger('validation'))
    result['originals_unchanged'] = all(digest(s.path.read_bytes()) == digest(s.content) for s in (before, after))
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
