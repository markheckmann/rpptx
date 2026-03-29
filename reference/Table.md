# Table proxy

Provides access to table cells, rows, columns, and formatting flags.
Access via `graphic_frame$table`.

## Methods

### Public methods

- [`Table$new()`](#method-Table-new)

- [`Table$cell()`](#method-Table-cell)

- [`Table$iter_cells()`](#method-Table-iter_cells)

- [`Table$clone()`](#method-Table-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    Table$new(tbl_elm, graphic_frame)

------------------------------------------------------------------------

### Method `cell()`

#### Usage

    Table$cell(row_idx, col_idx)

------------------------------------------------------------------------

### Method `iter_cells()`

#### Usage

    Table$iter_cells()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    Table$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
