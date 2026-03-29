# Alias: ChartData == CategoryChartData

Alias: ChartData == CategoryChartData

Alias: ChartData == CategoryChartData

## Super class

[`rpptx::BaseChartData`](https://markheckmann.github.io/rpptx/reference/BaseChartData.md)
-\> `CategoryChartData`

## Methods

### Public methods

- [`ChartData$add_series()`](#method-CategoryChartData-add_series)

- [`ChartData$add_category()`](#method-CategoryChartData-add_category)

- [`ChartData$categories_ref()`](#method-CategoryChartData-categories_ref)

- [`ChartData$values_ref()`](#method-CategoryChartData-values_ref)

- [`ChartData$.workbook_writer()`](#method-CategoryChartData-.workbook_writer)

- [`ChartData$clone()`](#method-CategoryChartData-clone)

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

    ChartData$add_series(name, values = numeric(0), number_format = NULL)

------------------------------------------------------------------------

### Method `add_category()`

#### Usage

    ChartData$add_category(label)

------------------------------------------------------------------------

### Method `categories_ref()`

#### Usage

    ChartData$categories_ref()

------------------------------------------------------------------------

### Method `values_ref()`

#### Usage

    ChartData$values_ref(series)

------------------------------------------------------------------------

### Method `.workbook_writer()`

#### Usage

    ChartData$.workbook_writer()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    ChartData$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
