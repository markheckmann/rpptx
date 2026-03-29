# Color format accessor

Wraps the "color choice parent" element (e.g. `<a:solidFill>`,
`<a:fgClr>`, `<a:bgClr>`, or a gradient stop `<a:gs>`) and exposes color
properties.

## Details

Access via `FillFormat$fore_color`, `FillFormat$back_color`, or
`LineFormat$color`.

## Methods

### Public methods

- [`ColorFormat$new()`](#method-ColorFormat-new)

- [`ColorFormat$clone()`](#method-ColorFormat-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    ColorFormat$new(xPr)

#### Arguments

- `xPr`:

  A ColorChoiceParent element (e.g. CT_SolidColorFillProperties).

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    ColorFormat$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
