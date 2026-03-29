# Chart data container for bubble charts

Holds series and bubble data points (x, y, bubble size). Use
`add_series()` to add a series, then call `add_data_point()` on each
returned series object.

## Super classes

`rpptx::BaseChartData` -\>
[`rpptx::XyChartData`](https://markheckmann.github.io/rpptx/reference/XyChartData.md)
-\> `BubbleChartData`

## Methods

### Public methods

- [`BubbleChartData$add_series()`](#method-BubbleChartData-add_series)

- [`BubbleChartData$bubble_sizes_ref()`](#method-BubbleChartData-bubble_sizes_ref)

- [`BubbleChartData$.workbook_writer()`](#method-BubbleChartData-.workbook_writer)

- [`BubbleChartData$clone()`](#method-BubbleChartData-clone)

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

    BubbleChartData$add_series(name, number_format = NULL)

------------------------------------------------------------------------

### Method `bubble_sizes_ref()`

#### Usage

    BubbleChartData$bubble_sizes_ref(series)

------------------------------------------------------------------------

### Method `.workbook_writer()`

#### Usage

    BubbleChartData$.workbook_writer()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    BubbleChartData$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
