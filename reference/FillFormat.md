# Fill format accessor

Wraps the shape-properties element (`<p:spPr>`) and exposes fill
properties. Access via `shape$fill`.

## Methods

### Public methods

- [`FillFormat$new()`](#method-FillFormat-new)

- [`FillFormat$solid()`](#method-FillFormat-solid)

- [`FillFormat$background()`](#method-FillFormat-background)

- [`FillFormat$gradient()`](#method-FillFormat-gradient)

- [`FillFormat$patterned()`](#method-FillFormat-patterned)

- [`FillFormat$clone()`](#method-FillFormat-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    FillFormat$new(spPr)

#### Arguments

- `spPr`:

  A CT_ShapeProperties or similar element that contains fill.

------------------------------------------------------------------------

### Method `solid()`

#### Usage

    FillFormat$solid()

------------------------------------------------------------------------

### Method `background()`

#### Usage

    FillFormat$background()

------------------------------------------------------------------------

### Method `gradient()`

#### Usage

    FillFormat$gradient()

------------------------------------------------------------------------

### Method `patterned()`

#### Usage

    FillFormat$patterned()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    FillFormat$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
