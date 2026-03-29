# Line format accessor

Wraps the shape-properties element (`<p:spPr>`) and exposes border/line
properties. Access via `shape$line`.

## Methods

### Public methods

- [`LineFormat$new()`](#method-LineFormat-new)

- [`LineFormat$clone()`](#method-LineFormat-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    LineFormat$new(spPr)

#### Arguments

- `spPr`:

  A CT_ShapeProperties element.

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    LineFormat$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
