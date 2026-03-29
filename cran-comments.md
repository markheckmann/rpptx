# CRAN submission comments — rpptx 0.1.0

## Test environments

* macOS 15 (arm64), R 4.5.x (local) — `devtools::check(cran = TRUE)`: Status OK
* GitHub Actions: ubuntu-latest (R devel, release, oldrel-1), windows-latest (release), macos-latest (release)

## R CMD check results

0 errors | 0 warnings | 0 notes

## Downstream dependencies

None — this is a new package.

## Notes for CRAN reviewers

* R port of the Python `python-pptx` library (MIT). No Python required at runtime.
* `inst/templates/` contains a default `.pptx` template and XML fragments used
  at runtime; these are not test fixtures.
* `tests/test_files/` contains `.pptx` fixture files for the read-path tests.
* All temporary `.pptx` files written during tests use `tempfile()` and are
  cleaned up automatically.
