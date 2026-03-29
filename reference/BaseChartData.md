# Base class for chart data objects

Base class for chart data objects

Base class for chart data objects

## Methods

### Public methods

- [`BaseChartData$new()`](#method-BaseChartData-new)

- [`BaseChartData$get()`](#method-BaseChartData-get)

- [`BaseChartData$to_list()`](#method-BaseChartData-to_list)

- [`BaseChartData$data_point_offset()`](#method-BaseChartData-data_point_offset)

- [`BaseChartData$series_index()`](#method-BaseChartData-series_index)

- [`BaseChartData$series_name_ref()`](#method-BaseChartData-series_name_ref)

- [`BaseChartData$x_values_ref()`](#method-BaseChartData-x_values_ref)

- [`BaseChartData$y_values_ref()`](#method-BaseChartData-y_values_ref)

- [`BaseChartData$xml_str()`](#method-BaseChartData-xml_str)

- [`BaseChartData$.workbook_writer()`](#method-BaseChartData-.workbook_writer)

- [`BaseChartData$clone()`](#method-BaseChartData-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    BaseChartData$new(number_format = "General")

------------------------------------------------------------------------

### Method [`get()`](https://rdrr.io/r/base/get.html)

#### Usage

    BaseChartData$get(i)

------------------------------------------------------------------------

### Method `to_list()`

#### Usage

    BaseChartData$to_list()

------------------------------------------------------------------------

### Method `data_point_offset()`

#### Usage

    BaseChartData$data_point_offset(series)

------------------------------------------------------------------------

### Method `series_index()`

#### Usage

    BaseChartData$series_index(series)

------------------------------------------------------------------------

### Method `series_name_ref()`

#### Usage

    BaseChartData$series_name_ref(series)

------------------------------------------------------------------------

### Method `x_values_ref()`

#### Usage

    BaseChartData$x_values_ref(series)

------------------------------------------------------------------------

### Method `y_values_ref()`

#### Usage

    BaseChartData$y_values_ref(series)

------------------------------------------------------------------------

### Method `xml_str()`

#### Usage

    BaseChartData$xml_str(chart_type)

------------------------------------------------------------------------

### Method `.workbook_writer()`

#### Usage

    BaseChartData$.workbook_writer()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    BaseChartData$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
