# Package index

## Open / save

- [`pptx_presentation()`](https://markheckmann.github.io/rpptx/reference/pptx_presentation.md)
  : Create or open a PowerPoint presentation

## Shapes

- [`FreeformBuilder`](https://markheckmann.github.io/rpptx/reference/FreeformBuilder.md)
  : Freeform shape builder

## Tables

- [`as.data.frame(`*`<Table>`*`)`](https://markheckmann.github.io/rpptx/reference/as.data.frame.Table.md)
  : Convert a Table to a data frame of cell text values

## Chart data

- [`CategoryChartData`](https://markheckmann.github.io/rpptx/reference/CategoryChartData.md)
  : Chart data container for category (bar, line, pie, etc.) charts
- [`XyChartData`](https://markheckmann.github.io/rpptx/reference/XyChartData.md)
  : Chart data container for XY (scatter) charts
- [`BubbleChartData`](https://markheckmann.github.io/rpptx/reference/BubbleChartData.md)
  : Chart data container for bubble charts

## Colors

- [`RGBColor()`](https://markheckmann.github.io/rpptx/reference/RGBColor.md)
  : Create an RGB color value
- [`RGBColor_from_str()`](https://markheckmann.github.io/rpptx/reference/RGBColor_from_str.md)
  : Create RGBColor from a 6-character hex string

## Units

- [`Length()`](https://markheckmann.github.io/rpptx/reference/Length.md)
  [`Inches()`](https://markheckmann.github.io/rpptx/reference/Length.md)
  [`Cm()`](https://markheckmann.github.io/rpptx/reference/Length.md)
  [`Mm()`](https://markheckmann.github.io/rpptx/reference/Length.md)
  [`Pt()`](https://markheckmann.github.io/rpptx/reference/Length.md)
  [`Emu()`](https://markheckmann.github.io/rpptx/reference/Length.md)
  [`Centipoints()`](https://markheckmann.github.io/rpptx/reference/Length.md)
  : Create a Length value in English Metric Units (EMU)
- [`as_inches()`](https://markheckmann.github.io/rpptx/reference/as_inches.md)
  [`as_cm()`](https://markheckmann.github.io/rpptx/reference/as_inches.md)
  [`as_mm()`](https://markheckmann.github.io/rpptx/reference/as_inches.md)
  [`as_pt()`](https://markheckmann.github.io/rpptx/reference/as_inches.md)
  [`as_emu()`](https://markheckmann.github.io/rpptx/reference/as_inches.md)
  [`as_centipoints()`](https://markheckmann.github.io/rpptx/reference/as_inches.md)
  : Convert a Length value to inches

## Shape enumerations

- [`MSO_AUTO_SHAPE_TYPE`](https://markheckmann.github.io/rpptx/reference/MSO_AUTO_SHAPE_TYPE.md)
  [`MSO_SHAPE`](https://markheckmann.github.io/rpptx/reference/MSO_AUTO_SHAPE_TYPE.md)
  : MSO_AUTO_SHAPE_TYPE — preset geometry names for autoshapes
- [`MSO_SHAPE_TYPE`](https://markheckmann.github.io/rpptx/reference/MSO_SHAPE_TYPE.md)
  : MSO_SHAPE_TYPE — shape type classification constants
- [`MSO_CONNECTOR_TYPE`](https://markheckmann.github.io/rpptx/reference/MSO_CONNECTOR_TYPE.md)
  [`MSO_CONNECTOR`](https://markheckmann.github.io/rpptx/reference/MSO_CONNECTOR_TYPE.md)
  : Connector type constants (MSO_CONNECTOR_TYPE)
- [`PP_PLACEHOLDER`](https://markheckmann.github.io/rpptx/reference/PP_PLACEHOLDER.md)
  : PP_PLACEHOLDER — placeholder type string constants

## Formatting enumerations

- [`MSO_THEME_COLOR`](https://markheckmann.github.io/rpptx/reference/MSO_THEME_COLOR.md)
  : MSO_THEME_COLOR — theme color XML value strings

- [`MSO_COLOR_TYPE`](https://markheckmann.github.io/rpptx/reference/MSO_COLOR_TYPE.md)
  : MSO_COLOR_TYPE — color source type strings

- [`MSO_FILL`](https://markheckmann.github.io/rpptx/reference/MSO_FILL.md)
  : MSO_FILL — fill type strings

- [`MSO_PATTERN_TYPE`](https://markheckmann.github.io/rpptx/reference/MSO_PATTERN_TYPE.md)
  : MSO_PATTERN_TYPE — preset pattern fill type strings

- [`MSO_LINE_DASH_STYLE`](https://markheckmann.github.io/rpptx/reference/MSO_LINE_DASH_STYLE.md)
  :

  MSO_LINE_DASH_STYLE — preset dash pattern strings for
  `<a:prstDash val="..."/>`

- [`MSO_AUTO_SIZE`](https://markheckmann.github.io/rpptx/reference/MSO_AUTO_SIZE.md)
  : MSO_AUTO_SIZE — text-box auto-sizing behaviour

- [`MSO_VERTICAL_ANCHOR`](https://markheckmann.github.io/rpptx/reference/MSO_VERTICAL_ANCHOR.md)
  : MSO_VERTICAL_ANCHOR — vertical text anchor

- [`PP_PARAGRAPH_ALIGNMENT`](https://markheckmann.github.io/rpptx/reference/PP_PARAGRAPH_ALIGNMENT.md)
  : PP_PARAGRAPH_ALIGNMENT — paragraph alignment

- [`PP_ACTION_TYPE`](https://markheckmann.github.io/rpptx/reference/PP_ACTION_TYPE.md)
  : PP_ACTION_TYPE — type of action for a click or hover

## Chart enumerations

- [`XL_CHART_TYPE`](https://markheckmann.github.io/rpptx/reference/XL_CHART_TYPE.md)
  : XL_CHART_TYPE — chart type enumeration
- [`XL_AXIS_CROSSES`](https://markheckmann.github.io/rpptx/reference/XL_AXIS_CROSSES.md)
  : XL_AXIS_CROSSES — where the other axis crosses this axis
- [`XL_CATEGORY_TYPE`](https://markheckmann.github.io/rpptx/reference/XL_CATEGORY_TYPE.md)
  : XL_CATEGORY_TYPE — category axis scale type
- [`XL_DATA_LABEL_POSITION`](https://markheckmann.github.io/rpptx/reference/XL_DATA_LABEL_POSITION.md)
  [`XL_LABEL_POSITION`](https://markheckmann.github.io/rpptx/reference/XL_DATA_LABEL_POSITION.md)
  : XL_DATA_LABEL_POSITION — position of a data label
- [`XL_LEGEND_POSITION`](https://markheckmann.github.io/rpptx/reference/XL_LEGEND_POSITION.md)
  : XL_LEGEND_POSITION — position of chart legend
- [`XL_MARKER_STYLE`](https://markheckmann.github.io/rpptx/reference/XL_MARKER_STYLE.md)
  : XL_MARKER_STYLE — shape of a data point marker
- [`XL_TICK_LABEL_POSITION`](https://markheckmann.github.io/rpptx/reference/XL_TICK_LABEL_POSITION.md)
  : XL_TICK_LABEL_POSITION — position of tick-mark labels on a chart
  axis
- [`XL_TICK_MARK`](https://markheckmann.github.io/rpptx/reference/XL_TICK_MARK.md)
  : XL_TICK_MARK — type of axis tick mark

## Utilities

- [`qn()`](https://markheckmann.github.io/rpptx/reference/qn.md) : Get
  the Clark-notation qualified tag name for a namespace-prefixed tag
