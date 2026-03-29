# Image value object

Immutable wrapper around raw image bytes that exposes ext, content_type,
sha1, size (pixels) and dpi.

## Methods

### Public methods

- [`Image$new()`](#method-Image-new)

- [`Image$sha1()`](#method-Image-sha1)

- [`Image$clone()`](#method-Image-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    Image$new(blob, filename = NULL)

------------------------------------------------------------------------

### Method `sha1()`

#### Usage

    Image$sha1()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    Image$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
