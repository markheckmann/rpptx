# Proxy wrapping an XML element with a parent reference

The parent is used to resolve the `part` property by delegation.

## Super class

[`rpptx::ElementProxy`](https://markheckmann.github.io/rpptx/reference/ElementProxy.md)
-\> `ParentedElementProxy`

## Methods

### Public methods

- [`ParentedElementProxy$new()`](#method-ParentedElementProxy-new)

- [`ParentedElementProxy$clone()`](#method-ParentedElementProxy-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    ParentedElementProxy$new(element, parent)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    ParentedElementProxy$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
