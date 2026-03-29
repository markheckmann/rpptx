# Return the appropriate chart XML writer for chart_type

Return the appropriate chart XML writer for chart_type

## Usage

``` r
chart_xml_writer(chart_type, chart_data)
```

## Arguments

- chart_type:

  An integer from XL_CHART_TYPE, e.g. `XL_CHART_TYPE$COLUMN_CLUSTERED`.

- chart_data:

  A
  [CategoryChartData](https://markheckmann.github.io/rpptx/reference/CategoryChartData.md),
  [XyChartData](https://markheckmann.github.io/rpptx/reference/XyChartData.md),
  or
  [BubbleChartData](https://markheckmann.github.io/rpptx/reference/BubbleChartData.md)
  object.

## Value

An R6 writer object with an `$xml` active binding returning the XML
string.
