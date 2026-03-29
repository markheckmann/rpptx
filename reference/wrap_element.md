# Wrap an xml2 node in the appropriate R6 element class

Looks up the node's tag in the element registry and returns an instance
of the registered class, or a BaseOxmlElement if no class is registered.

## Usage

``` r
wrap_element(node)
```

## Arguments

- node:

  An xml2 xml_node.

## Value

A BaseOxmlElement (or subclass) instance.
