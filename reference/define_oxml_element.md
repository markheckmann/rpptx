# Define a custom XML element R6 class

Factory function that generates an R6 class with active bindings for
child elements and attributes. This is the R equivalent of python-pptx's
MetaOxmlElement metaclass + xmlchemy descriptors.

## Usage

``` r
define_oxml_element(
  classname,
  tag,
  children = list(),
  attributes = list(),
  methods = list(),
  active = list(),
  inherit = BaseOxmlElement
)
```

## Arguments

- classname:

  Name for the R6 class.

- tag:

  Namespace-prefixed tag, e.g. `"p:presentation"`.

- children:

  Named list of child element specs created with
  [`zero_or_one()`](https://markheckmann.github.io/rpptx/reference/zero_or_one.md),
  [`zero_or_more()`](https://markheckmann.github.io/rpptx/reference/zero_or_more.md),
  [`one_or_more()`](https://markheckmann.github.io/rpptx/reference/one_or_more.md),
  or
  [`one_and_only_one()`](https://markheckmann.github.io/rpptx/reference/one_and_only_one.md).

- attributes:

  Named list of attribute specs created with
  [`optional_attribute()`](https://markheckmann.github.io/rpptx/reference/optional_attribute.md)
  or
  [`required_attribute()`](https://markheckmann.github.io/rpptx/reference/required_attribute.md).

- methods:

  Named list of additional public methods (as functions).

- active:

  Named list of additional active bindings.

- inherit:

  R6 class to inherit from (default: BaseOxmlElement).

## Value

An R6ClassGenerator.
