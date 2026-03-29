# Picture crop descriptor

Provides read/write access to the per-edge crop fractions (0.0–1.0) of a
picture. A value of 0.1 means 10% of the edge is cropped. Access via
`picture$crop`.

## Methods

### Public methods

- [`PictureCrop$new()`](#method-PictureCrop-new)

- [`PictureCrop$clone()`](#method-PictureCrop-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    PictureCrop$new(blipFill)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    PictureCrop$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
