# Define an OptionalAttribute specification

Define an OptionalAttribute specification

## Usage

``` r
optional_attribute(attr_name, simple_type, default = NULL)
```

## Arguments

- attr_name:

  Attribute name (may be namespace-prefixed).

- simple_type:

  A list with `from_xml` and `to_xml` functions.

- default:

  Default value when attribute is absent (default NULL).

## Value

A list describing the attribute spec.
