# Create a new XML element with the given namespace-prefixed tag

Create a new XML element with the given namespace-prefixed tag

## Usage

``` r
OxmlElement(nsptag_str, nsmap = NULL)
```

## Arguments

- nsptag_str:

  Namespace-prefixed tag, e.g. `"a:tbl"`.

- nsmap:

  Optional named character vector of additional namespace mappings.

## Value

A wrapped BaseOxmlElement (or appropriate subclass).
