"""Compare old and fixed fixture cleanup with weakrefs; own Tk windows only.

This proves fixture lifetime/affinity, not the CI's exact random GC schedule.
"""
import gc
import json
from pathlib import Path
import sys
import threading
import weakref

ROOT = Path(__file__).resolve().parents[4]
sys.path[:0] = [str(ROOT), str(ROOT / "tests")]
from test_windows_gui_helpers import WindowTests


def observe_cleanup(use_fixed):
    case = WindowTests("test_result_directory_unavailable_until_output_exists")
    case.setUp()
    finalized = []
    references = {}
    resources = {"app": case.app, "root": case.root,
                 "status": case.app.status, "all_merges": case.app.all_merges}
    for name, resource in resources.items():
        references[name] = weakref.ref(resource)
        weakref.finalize(resource, lambda label=name: finalized.append(
            {"object": label, "thread": threading.current_thread().name}))
    del resources, resource
    try:
        if use_fixed:
            case.tearDown()
        else:
            # Exact previous fixture behavior: widgets destroyed, attributes kept.
            if case.root is not None:
                case.root.destroy()
            case.directory.cleanup()
        retained = {name: reference() is not None for name, reference in references.items()}
    finally:
        # Always clean the baseline probe on the creating thread too.
        case.app = None
        case.root = None
        gc.collect()
    assert len(finalized) == 4 and all(item["thread"] == "MainThread" for item in finalized)
    return {"retained_after_teardown": retained, "finalization": finalized}


def main():
    baseline, fixed = observe_cleanup(False), observe_cleanup(True)
    assert all(baseline["retained_after_teardown"].values())
    assert not any(fixed["retained_after_teardown"].values())
    print(json.dumps({"baseline": baseline, "fixed": fixed, "passed": True},
                     ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
