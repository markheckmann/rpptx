# Chart data container for category (bar, line, pie, etc.) charts

Holds categories and one or more series of values. Use `add_series()` to
add data series. Set `categories` by assigning a character/numeric/Date
vector to `$categories`.

## Super class

[`rpptx::BaseChartData`](https://markheckmann.github.io/rpptx/reference/BaseChartData.md)
-\> `CategoryChartData`

## Methods

### Public methods

- [`CategoryChartData$add_series()`](#method-CategoryChartData-add_series)

- [`CategoryChartData$add_category()`](#method-CategoryChartData-add_category)

- [`CategoryChartData$categories_ref()`](#method-CategoryChartData-categories_ref)

- [`CategoryChartData$values_ref()`](#method-CategoryChartData-values_ref)

- [`CategoryChartData$.workbook_writer()`](#method-CategoryChartData-.workbook_writer)

- [`CategoryChartData$clone()`](#method-CategoryChartData-clone)

Inherited methods

- [`rpptx::BaseChartData$data_point_offset()`](https://markheckmann.github.io/rpptx/reference/BaseChartData.html#method-data_point_offset)
- [`rpptx::BaseChartData$get()`](https://markheckmann.github.io/rpptx/reference/BaseChartData.html#method-get)
- [`rpptx::BaseChartData$initialize()`](https://markheckmann.github.io/rpptx/reference/BaseChartData.html#method-initialize)
- [`rpptx::BaseChartData$series_index()`](https://markheckmann.github.io/rpptx/reference/BaseChartData.html#method-series_index)
- [`rpptx::BaseChartData$series_name_ref()`](https://markheckmann.github.io/rpptx/reference/BaseChartData.html#method-series_name_ref)
- [`rpptx::BaseChartData$to_list()`](https://markheckmann.github.io/rpptx/reference/BaseChartData.html#method-to_list)
- [`rpptx::BaseChartData$x_values_ref()`](https://markheckmann.github.io/rpptx/reference/BaseChartData.html#method-x_values_ref)
- [`rpptx::BaseChartData$xml_str()`](https://markheckmann.github.io/rpptx/reference/BaseChartData.html#method-xml_str)
- [`rpptx::BaseChartData$y_values_ref()`](https://markheckmann.github.io/rpptx/reference/BaseChartData.html#method-y_values_ref)

------------------------------------------------------------------------

### Method `add_series()`

#### Usage

    CategoryChartData$add_series(name, values = numeric(0), number_format = NULL)

------------------------------------------------------------------------

### Method `add_category()`

#### Usage

    CategoryChartData$add_category(label)

------------------------------------------------------------------------

### Method `categories_ref()`

#### Usage

    CategoryChartData$categories_ref()

------------------------------------------------------------------------

### Method `values_ref()`

#### Usage

    CategoryChartData$values_ref(series)

------------------------------------------------------------------------

### Method `.workbook_writer()`

#### Usage

    CategoryChartData$.workbook_writer()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    CategoryChartData$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.

## Examples

``` r
cd <- CategoryChartData$new()
cd$categories <- c("Q1", "Q2", "Q3")
cd$add_series("Sales", c(100, 200, 150))
```
