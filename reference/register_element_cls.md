# Register an R6 class for a given XML element tag

When the XML parser encounters an element with this tag, it will be
wrapped in an instance of `cls` instead of the default BaseOxmlElement.

## Usage

``` r
register_element_cls(nsptagname, cls)
```

## Arguments

- nsptagname:

  Namespace-prefixed tag, e.g. `"p:presentation"`.

- cls:

  An R6ClassGenerator (e.g. the result of
  [`define_oxml_element()`](https://markheckmann.github.io/rpptx/reference/define_oxml_element.md)).
