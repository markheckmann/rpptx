# Character properties for a run, paragraph-default, or end-of-paragraph.

Wraps an `a:rPr`, `a:defRPr`, or `a:endParaRPr` element. Provides access
to font name, size, bold, italic, and underline. All properties are R/W;
assigning NULL removes the override and inherits from the style
hierarchy.

## Methods

### Public methods

- [`Font$new()`](#method-Font-new)

- [`Font$element()`](#method-Font-element)

- [`Font$clone()`](#method-Font-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    Font$new(rPr, run = NULL)

------------------------------------------------------------------------

### Method `element()`

#### Usage

    Font$element()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    Font$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
