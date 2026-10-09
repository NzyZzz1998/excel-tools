# CI temporary-directory path alias correction

Date: 2026-10-10. Failed run: `37956290587`.

The Windows CI job ran all 115 tests. Its sole failure was the directory equality assertion in `test_cleanup_failure_preserves_original_errno_and_reports_own_part`: the engine's resolved path used `C:/Users/runneradmin/...`, while `TemporaryDirectory` retained the equivalent `C:/Users/RUNNER~1/...` spelling. GUI tests passed in that run. This was a test identity comparison failure, not an output-cleanup failure.

The assertion now uses `error.partial_path.parent.samefile(self.root)` to verify that both paths identify the same existing directory. The original exception identity, errno, actual partial write, cleanup exception, `.part` existence, diagnostic path, absence of a final workbook, and unchanged source assertions remain intact.

The entire test file was reviewed for similar comparisons. Other full-path comparisons either derive both operands from the same temporary-root spelling or compare paths returned by the same engine; filename checks deliberately test naming. The direct `Path.unlink` fault-injection comparison receives the same path passed to `_publish_part`, without the public API's path resolution. No further change was necessary.

Local targeted verification:

```powershell
$env:PYTHONPATH = (Join-Path $env:TEMP 'excel-tools-review-20261008-deps') + ';' + (Join-Path (Get-Location) 'tests')
python -m unittest test_engine_io_progress -v
```

Result: **14 tests passed**, 0.390 seconds, exit 0. The exact runner path-alias environment still requires the next CI run; this local run does not claim to reproduce that environment.

Only the test assertion changed. Application identities remain:

- Engine SHA-256: `602f867dfd9be498e911a6599cd8da1ffcdaed5dd6bb20db5af9637b0dc196f3`.
- GUI SHA-256: `50ea03ad60ed6f86003132b9d4c951213f5f01b6c202a677bf765a8df01d2094`.

No application rebuild or business workbook rerun was performed.
