# Load an XmlPart from a blob

Load an XmlPart from a blob

## Usage

``` r
XmlPart_load(cls, partname, content_type, package, blob)
```

## Arguments

- cls:

  The R6 class generator to use.

- partname:

  A PackURI.

- content_type:

  Content type string.

- package:

  The package.

- blob:

  Raw bytes of XML.

## Value

An XmlPart instance.
