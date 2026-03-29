# Video value object

Wraps a video bytestream and its MIME type. Used by
`SlideShapes$add_movie()`.

## Methods

### Public methods

- [`Video$new()`](#method-Video-new)

- [`Video$clone()`](#method-Video-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    Video$new(blob, mime_type = NULL, filename = NULL)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    Video$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
