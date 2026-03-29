# Chart data container for XY (scatter) charts

Chart data container for XY (scatter) charts

Chart data container for XY (scatter) charts

## Super class

[`rpptx::BaseChartData`](https://markheckmann.github.io/rpptx/reference/BaseChartData.md)
-\> `XyChartData`

## Methods

### Public methods

- [`XyChartData$add_series()`](#method-XyChartData-add_series)

- [`XyChartData$.workbook_writer()`](#method-XyChartData-.workbook_writer)

- [`XyChartData$clone()`](#method-XyChartData-clone)

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

    XyChartData$add_series(
      name,
      x_values = NULL,
      y_values = NULL,
      number_format = NULL
    )

------------------------------------------------------------------------

### Method `.workbook_writer()`

#### Usage

    XyChartData$.workbook_writer()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    XyChartData$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
