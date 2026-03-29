# Get the Clark name of an xml2 node

xml2 nodes can report their tag in various ways. This function returns
the Clark-notation tag `{uri}localname` for use in the element registry.

## Usage

``` r
xml_clark_name(node)
```

## Arguments

- node:

  An xml2 xml_node.

## Value

A string in Clark notation, or the plain tag name if no namespace.
