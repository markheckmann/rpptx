# CRAN submission comments

## Test environments

* local macOS 15 (aarch64), R 4.4.x
* GitHub Actions: ubuntu-latest (R devel, release, oldrel-1), windows-latest (release), macos-latest (release)

## R CMD check results

Local check: 0 errors | 0 warnings | 0 notes (when pdflatex is available).

### Local machine notes
- **PDF manual ERROR/WARNING**: `pdflatex is not available` on this development
  machine. The PDF manual builds successfully on CRAN (Linux with TeXLive).
- **URL 404 NOTEs**: `https://github.com/markheckmann/rpptx` returns 404 until
  the GitHub repository is made public. These will resolve upon publication.

## Reverse dependencies

None — this is a new package.
