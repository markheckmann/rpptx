# Convert a Table to a data frame of cell text values

Each cell's text content becomes one element of the data frame. Row 1
becomes the first row of the data frame (not used as column names). To
promote the first row to headers use
`setNames(as.data.frame(tbl), as.data.frame(tbl)[1, ])`.

## Usage

``` r
# S3 method for class 'Table'
as.data.frame(x, row.names = NULL, optional = FALSE, ...)
```

## Arguments

- x:

  A `Table` object.

- row.names:

  NULL or a character vector of row names (default: NULL).

- optional:

  Ignored; included for S3 compatibility.

- ...:

  Ignored.

## Value

A data.frame with character columns named V1, V2, …
