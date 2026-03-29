# Custom element class for p:spTree and p:grpSp elements

Custom element class for p:spTree and p:grpSp elements

Custom element class for p:spTree and p:grpSp elements

## Super classes

[`rpptx::BaseOxmlElement`](https://markheckmann.github.io/rpptx/reference/BaseOxmlElement.md)
-\>
[`rpptx::BaseShapeElement`](https://markheckmann.github.io/rpptx/reference/BaseShapeElement.md)
-\> `CT_GroupShape`

## Methods

### Public methods

- [`CT_GroupShape$iter_shape_elms()`](#method-CT_GroupShape-iter_shape_elms)

- [`CT_GroupShape$iter_ph_elms()`](#method-CT_GroupShape-iter_ph_elms)

- [`CT_GroupShape$max_shape_id()`](#method-CT_GroupShape-max_shape_id)

- [`CT_GroupShape$next_shape_id()`](#method-CT_GroupShape-next_shape_id)

- [`CT_GroupShape$add_textbox()`](#method-CT_GroupShape-add_textbox)

- [`CT_GroupShape$add_autoshape()`](#method-CT_GroupShape-add_autoshape)

- [`CT_GroupShape$add_placeholder()`](#method-CT_GroupShape-add_placeholder)

- [`CT_GroupShape$add_table()`](#method-CT_GroupShape-add_table)

- [`CT_GroupShape$add_chart()`](#method-CT_GroupShape-add_chart)

- [`CT_GroupShape$add_pic()`](#method-CT_GroupShape-add_pic)

- [`CT_GroupShape$add_cxnSp()`](#method-CT_GroupShape-add_cxnSp)

- [`CT_GroupShape$add_freeform_sp()`](#method-CT_GroupShape-add_freeform_sp)

- [`CT_GroupShape$add_grpSp()`](#method-CT_GroupShape-add_grpSp)

- [`CT_GroupShape$clone()`](#method-CT_GroupShape-clone)

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
- [`rpptx::BaseShapeElement$get_or_add_ln()`](https://markheckmann.github.io/rpptx/reference/BaseShapeElement.html#method-get_or_add_ln)
- [`rpptx::BaseShapeElement$get_or_add_xfrm()`](https://markheckmann.github.io/rpptx/reference/BaseShapeElement.html#method-get_or_add_xfrm)

------------------------------------------------------------------------

### Method `iter_shape_elms()`

#### Usage

    CT_GroupShape$iter_shape_elms()

------------------------------------------------------------------------

### Method `iter_ph_elms()`

#### Usage

    CT_GroupShape$iter_ph_elms()

------------------------------------------------------------------------

### Method `max_shape_id()`

#### Usage

    CT_GroupShape$max_shape_id()

------------------------------------------------------------------------

### Method `next_shape_id()`

#### Usage

    CT_GroupShape$next_shape_id()

------------------------------------------------------------------------

### Method `add_textbox()`

#### Usage

    CT_GroupShape$add_textbox(id, name, x, y, cx, cy)

------------------------------------------------------------------------

### Method `add_autoshape()`

#### Usage

    CT_GroupShape$add_autoshape(id, name, prst, x, y, cx, cy)

------------------------------------------------------------------------

### Method `add_placeholder()`

#### Usage

    CT_GroupShape$add_placeholder(id, name, ph_type, orient, sz, idx)

------------------------------------------------------------------------

### Method `add_table()`

#### Usage

    CT_GroupShape$add_table(id, name, rows, cols, x, y, cx, cy)

------------------------------------------------------------------------

### Method `add_chart()`

#### Usage

    CT_GroupShape$add_chart(id, name, rId, x, y, cx, cy)

------------------------------------------------------------------------

### Method `add_pic()`

#### Usage

    CT_GroupShape$add_pic(id, name, desc, rId, x, y, cx, cy)

------------------------------------------------------------------------

### Method `add_cxnSp()`

#### Usage

    CT_GroupShape$add_cxnSp(
      id,
      name,
      prst,
      x,
      y,
      cx,
      cy,
      flipH = FALSE,
      flipV = FALSE
    )

------------------------------------------------------------------------

### Method `add_freeform_sp()`

#### Usage

    CT_GroupShape$add_freeform_sp(x, y, cx, cy)

------------------------------------------------------------------------

### Method `add_grpSp()`

#### Usage

    CT_GroupShape$add_grpSp(id, name, shape_elms)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    CT_GroupShape$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
