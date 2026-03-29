# Get the integer index from an array-style PackURI

Returns the trailing integer from the filename, e.g. 21 for
`"/ppt/slides/slide21.xml"`, or NULL for singleton parts.

## Usage

``` r
pack_uri_idx(uri)
```

## Arguments

- uri:

  A PackURI.

## Value

An integer or NULL.
