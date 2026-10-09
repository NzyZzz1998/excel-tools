# v1.1.1 GUI / engine integration review

Reviewed on 2026-10-09. This is an independent, read-only source and test review; it did not rerun the test suite, launch the GUI, or interact with a running user application.

## Reviewed source identities

| File | SHA-256 |
| --- | --- |
| `excel_unmerge_fill.py` | `602f867dfd9be498e911a6599cd8da1ffcdaed5dd6bb20db5af9637b0dc196f3` |
| `excel_unmerge_gui.py` | `50ea03ad60ed6f86003132b9d4c951213f5f01b6c202a677bf765a8df01d2094` |
| `tests/test_gui_batch.py` | `c429cf20dc3eb65721de7174cb3d6dd41f4a75a34d8ada637375e1b286df07e4` |
| `tests/test_windows_gui_helpers.py` | `1ceb49b6828cdf5ae31a9a63ea7f19b6eeec2f13519e970a77e498dbce674da9` |

## Findings

No reachable blocking regression was identified in the reviewed integration scope.

- **Event delivery and completion:** `process_batch` copies each synchronous engine progress dictionary into the queue between that file's `started` and `result` events. The worker does not access Tk. The GUI accepts progress only for the active file while running; late progress after a result or `done` cannot replace terminal status. Only results advance the file completion count and progress bar. A completed saving phase therefore does not prematurely report publication success.
- **Stop behavior:** stop remains a request to finish the current file safely and stop before the next file. Status refresh and result handling respect the stop flag, so later engine progress does not overwrite the waiting message. The window remains alive until processing finishes; there is no new mid-file cancellation path.
- **Elapsed time:** the displayed duration uses the monotonic clock, begins separately for each consumed `started` event, refreshes through the regular Tk polling loop, and clears on result or completion. Phase counters are labelled as processed rows or bytes; unknown totals do not produce a percentage or ETA. The timer is elapsed time, not evidence that a blocked underlying operation is making progress.
- **Recovery target:** selecting B preserves the previous task's failed or pending A entries. Recovery explicitly displays and processes A using its original merge option, while retaining B as the next start selection. On recovery completion the B list returns; starting a new batch resets prior task state and uses B. Recovery also visibly restores A's option, so a different option for the next new batch must be selected before starting it.
- **Regression test intent:** filtering progress events in existing batch assertions retains the original terminal event assertions. New tests check dictionary snapshots, file ordering, elapsed refresh without a new event, per-file timer reset, late progress after failures, stop-message priority, both recovery entry points, and minimum-window layout. Existing stop-at-boundary, captured-option, and thread-start-failure assertions remain present. This review makes no claim of additional execution beyond the separate acceptance test logs.

## Limits

This conclusion covers the source identities above and the specified integration paths. It does not independently repeat packaged-EXE operation, large-workbook performance, Windows file-lock fault injection, or physical display testing. Those remain supported by their separate release acceptance evidence. Ordinary progress observer exceptions are isolated by the engine; `BaseException` propagation remains the documented internal API boundary.
