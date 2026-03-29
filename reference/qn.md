# Get the Clark-notation qualified tag name for a namespace-prefixed tag

Converts a namespace-prefixed tag like `"p:cSld"` to Clark notation like
`"{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"`.

## Usage

``` r
qn(nsptag)
```

## Arguments

- nsptag:

  A namespace-prefixed tag string, e.g. `"p:cSld"`.

## Value

A string in Clark notation.

## Examples

``` r
qn("p:cSld")
#> [1] "{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"
qn("a:r")
#> [1] "{http://schemas.openxmlformats.org/drawingml/2006/main}r"
```
