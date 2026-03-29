# Table cell proxy

Provides access to cell text, fill, margins, and merge state. Access via
`table$cell(row, col)` or `table$rows[[i]]$cells[[j]]`.

## Methods

### Public methods

- [`TableCell$new()`](#method-TableCell-new)

- [`TableCell$merge()`](#method-TableCell-merge)

- [`TableCell$clone()`](#method-TableCell-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    TableCell$new(tc, parent)

------------------------------------------------------------------------

### Method [`merge()`](https://rdrr.io/r/base/merge.html)

#### Usage

    TableCell$merge(other_cell)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    TableCell$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
