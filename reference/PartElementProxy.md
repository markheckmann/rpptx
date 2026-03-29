# Proxy wrapping a part's root XML element

Used for domain objects that correspond directly to an OPC part (e.g.
Presentation wraps the root element of PresentationPart).

## Super class

[`rpptx::ElementProxy`](https://markheckmann.github.io/rpptx/reference/ElementProxy.md)
-\> `PartElementProxy`

## Methods

### Public methods

- [`PartElementProxy$new()`](#method-PartElementProxy-new)

- [`PartElementProxy$clone()`](#method-PartElementProxy-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    PartElementProxy$new(element, part)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    PartElementProxy$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
