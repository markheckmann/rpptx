# Value axis XML element

Value axis XML element

Value axis XML element

## Super classes

[`rpptx::BaseOxmlElement`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.md)
-\>
[`rpptx::BaseAxisElement`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.md)
-\> `CT_ValAx`

## Methods

### Public methods

- [`CT_ValAx$.remove_crosses()`](#method-CT_ValAx-.remove_crosses)

- [`CT_ValAx$.remove_crossesAt()`](#method-CT_ValAx-.remove_crossesAt)

- [`CT_ValAx$.add_crosses()`](#method-CT_ValAx-.add_crosses)

- [`CT_ValAx$.add_crossesAt()`](#method-CT_ValAx-.add_crossesAt)

- [`CT_ValAx$.add_majorUnit()`](#method-CT_ValAx-.add_majorUnit)

- [`CT_ValAx$.remove_majorUnit()`](#method-CT_ValAx-.remove_majorUnit)

- [`CT_ValAx$.add_minorUnit()`](#method-CT_ValAx-.add_minorUnit)

- [`CT_ValAx$.remove_minorUnit()`](#method-CT_ValAx-.remove_minorUnit)

- [`CT_ValAx$clone()`](#method-CT_ValAx-clone)

Inherited methods

- [`rpptx::BaseOxmlElement$append_child()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-append_child)
- [`rpptx::BaseOxmlElement$find()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-find)
- [`rpptx::BaseOxmlElement$findall()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-findall)
- [`rpptx::BaseOxmlElement$first_child_found_in()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-first_child_found_in)
- [`rpptx::BaseOxmlElement$get_attr()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-get_attr)
- [`rpptx::BaseOxmlElement$get_attrs()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-get_attrs)
- [`rpptx::BaseOxmlElement$get_node()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-get_node)
- [`rpptx::BaseOxmlElement$initialize()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-initialize)
- [`rpptx::BaseOxmlElement$insert_element_before()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-insert_element_before)
- [`rpptx::BaseOxmlElement$remove_all()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-remove_all)
- [`rpptx::BaseOxmlElement$remove_attr()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-remove_attr)
- [`rpptx::BaseOxmlElement$remove_child()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-remove_child)
- [`rpptx::BaseOxmlElement$set_attr()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-set_attr)
- [`rpptx::BaseOxmlElement$to_xml()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-to_xml)
- [`rpptx::BaseOxmlElement$xpath()`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.html#method-xpath)
- [`rpptx::BaseAxisElement$.add_majorTickMark()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-.add_majorTickMark)
- [`rpptx::BaseAxisElement$.add_minorTickMark()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-.add_minorTickMark)
- [`rpptx::BaseAxisElement$.remove_majorGridlines()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-.remove_majorGridlines)
- [`rpptx::BaseAxisElement$.remove_majorTickMark()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-.remove_majorTickMark)
- [`rpptx::BaseAxisElement$.remove_minorGridlines()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-.remove_minorGridlines)
- [`rpptx::BaseAxisElement$.remove_minorTickMark()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-.remove_minorTickMark)
- [`rpptx::BaseAxisElement$.remove_title()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-.remove_title)
- [`rpptx::BaseAxisElement$get_or_add_delete_()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-get_or_add_delete_)
- [`rpptx::BaseAxisElement$get_or_add_majorGridlines()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-get_or_add_majorGridlines)
- [`rpptx::BaseAxisElement$get_or_add_minorGridlines()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-get_or_add_minorGridlines)
- [`rpptx::BaseAxisElement$get_or_add_numFmt()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-get_or_add_numFmt)
- [`rpptx::BaseAxisElement$get_or_add_scaling()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-get_or_add_scaling)
- [`rpptx::BaseAxisElement$get_or_add_spPr()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-get_or_add_spPr)
- [`rpptx::BaseAxisElement$get_or_add_tickLblPos()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-get_or_add_tickLblPos)
- [`rpptx::BaseAxisElement$get_or_add_title()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-get_or_add_title)
- [`rpptx::BaseAxisElement$get_or_add_txPr()`](https://markheckmann.github.io/rpptx/reference/BaseAxisElement.html#method-get_or_add_txPr)

------------------------------------------------------------------------

### Method `.remove_crosses()`

#### Usage

    CT_ValAx$.remove_crosses()

------------------------------------------------------------------------

### Method `.remove_crossesAt()`

#### Usage

    CT_ValAx$.remove_crossesAt()

------------------------------------------------------------------------

### Method `.add_crosses()`

#### Usage

    CT_ValAx$.add_crosses(val = NULL)

------------------------------------------------------------------------

### Method `.add_crossesAt()`

#### Usage

    CT_ValAx$.add_crossesAt(val = NULL)

------------------------------------------------------------------------

### Method `.add_majorUnit()`

#### Usage

    CT_ValAx$.add_majorUnit(val = NULL)

------------------------------------------------------------------------

### Method `.remove_majorUnit()`

#### Usage

    CT_ValAx$.remove_majorUnit()

------------------------------------------------------------------------

### Method `.add_minorUnit()`

#### Usage

    CT_ValAx$.add_minorUnit(val = NULL)

------------------------------------------------------------------------

### Method `.remove_minorUnit()`

#### Usage

    CT_ValAx$.remove_minorUnit()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    CT_ValAx$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
