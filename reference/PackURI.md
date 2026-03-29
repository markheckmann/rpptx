# Create a PackURI (OPC part name)

A PackURI is a string that represents an absolute URI within an OPC
package. It must begin with a forward slash.

## Usage

``` r
PackURI(uri_str)
```

## Arguments

- uri_str:

  A string starting with "/", e.g. `"/ppt/slides/slide1.xml"`.

## Value

A PackURI (character with S3 class).
