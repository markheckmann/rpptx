# Package index

## Open / save

- [`pptx_presentation()`](https://markheckmann.github.io/rpptx/reference/pptx_presentation.md)
  : Create or open a PowerPoint presentation

## Presentation

- [`Presentation`](https://markheckmann.github.io/rpptx/reference/Presentation.md)
  : Presentation object

## Slides

- [`Slide`](https://markheckmann.github.io/rpptx/reference/Slide.md) :
  Slide object
- [`SlideLayout`](https://markheckmann.github.io/rpptx/reference/SlideLayout.md)
  : Slide layout object
- [`SlideMaster`](https://markheckmann.github.io/rpptx/reference/SlideMaster.md)
  : Slide master object
- [`Slides`](https://markheckmann.github.io/rpptx/reference/Slides.md) :
  Sequence of slides in a presentation
- [`SlideLayouts`](https://markheckmann.github.io/rpptx/reference/SlideLayouts.md)
  : Sequence of slide layouts for a slide master
- [`SlideMasters`](https://markheckmann.github.io/rpptx/reference/SlideMasters.md)
  : Sequence of slide masters in a presentation

## Shapes — collections

- [`SlideShapes`](https://markheckmann.github.io/rpptx/reference/SlideShapes.md)
  : Shape collection for a slide

## Shapes — proxies

- [`BaseShape`](https://markheckmann.github.io/rpptx/reference/BaseShape.md)
  : Base class for shape proxy objects
- [`Shape`](https://markheckmann.github.io/rpptx/reference/Shape.md) :
  Shape proxy for p:sp elements
- [`Picture`](https://markheckmann.github.io/rpptx/reference/Picture.md)
  : Shape proxy for p:pic elements
- [`GraphicFrame`](https://markheckmann.github.io/rpptx/reference/GraphicFrame.md)
  : Shape proxy for p:graphicFrame elements
- [`GroupShape`](https://markheckmann.github.io/rpptx/reference/GroupShape.md)
  : Shape proxy for p:grpSp (group shape) elements
- [`Connector`](https://markheckmann.github.io/rpptx/reference/Connector.md)
  : Shape proxy for p:cxnSp elements
- [`SlidePlaceholder`](https://markheckmann.github.io/rpptx/reference/SlidePlaceholder.md)
  : Placeholder shape on a slide
- [`LayoutPlaceholder`](https://markheckmann.github.io/rpptx/reference/LayoutPlaceholder.md)
  : Placeholder shape on a slide layout
- [`MasterPlaceholder`](https://markheckmann.github.io/rpptx/reference/MasterPlaceholder.md)
  : Placeholder shape on a slide master

## Freeform shapes

- [`FreeformBuilder`](https://markheckmann.github.io/rpptx/reference/FreeformBuilder.md)
  : Freeform shape builder

## Text

- [`TextFrame`](https://markheckmann.github.io/rpptx/reference/TextFrame.md)
  :

  Text frame corresponding to a `p:txBody` element.

- [`Paragraph`](https://markheckmann.github.io/rpptx/reference/Paragraph.md)
  :

  Paragraph object corresponding to an `a:p` element.

- [`Run`](https://markheckmann.github.io/rpptx/reference/Run.md) :

  Text run object corresponding to an `a:r` element.

- [`Font`](https://markheckmann.github.io/rpptx/reference/Font.md) :
  Character properties for a run, paragraph-default, or
  end-of-paragraph.

## Tables

- [`Table`](https://markheckmann.github.io/rpptx/reference/Table.md) :
  Table proxy
- [`TableCell`](https://markheckmann.github.io/rpptx/reference/TableCell.md)
  : Table cell proxy
- [`as.data.frame(`*`<Table>`*`)`](https://markheckmann.github.io/rpptx/reference/as.data.frame.Table.md)
  : Convert a Table to a data frame of cell text values

## Charts

- [`CategoryChartData`](https://markheckmann.github.io/rpptx/reference/CategoryChartData.md)
  : Chart data container for category (bar, line, pie, etc.) charts
- [`ChartData`](https://markheckmann.github.io/rpptx/reference/ChartData.md)
  : Alias: ChartData == CategoryChartData
- [`XyChartData`](https://markheckmann.github.io/rpptx/reference/XyChartData.md)
  : Chart data container for XY (scatter) charts
- [`BubbleChartData`](https://markheckmann.github.io/rpptx/reference/BubbleChartData.md)
  : Chart data container for bubble charts

## Formatting

- [`FillFormat`](https://markheckmann.github.io/rpptx/reference/FillFormat.md)
  : Fill format accessor
- [`GradientStop`](https://markheckmann.github.io/rpptx/reference/GradientStop.md)
  : Gradient stop proxy
- [`GradientStops`](https://markheckmann.github.io/rpptx/reference/GradientStops.md)
  : Gradient stops collection
- [`LineFormat`](https://markheckmann.github.io/rpptx/reference/LineFormat.md)
  : Line format accessor
- [`ColorFormat`](https://markheckmann.github.io/rpptx/reference/ColorFormat.md)
  : Color format accessor
- [`RGBColor()`](https://markheckmann.github.io/rpptx/reference/RGBColor.md)
  : Create an RGB color value
- [`ShadowFormat`](https://markheckmann.github.io/rpptx/reference/ShadowFormat.md)
  : Shape shadow effect

## Actions / hyperlinks

- [`ActionSetting`](https://markheckmann.github.io/rpptx/reference/ActionSetting.md)
  : Mouse action settings on a shape
- [`ShapeHyperlink`](https://markheckmann.github.io/rpptx/reference/ShapeHyperlink.md)
  : Hyperlink on a shape

## Enumerations

- [`MSO_AUTO_SHAPE_TYPE`](https://markheckmann.github.io/rpptx/reference/MSO_AUTO_SHAPE_TYPE.md)
  [`MSO_SHAPE`](https://markheckmann.github.io/rpptx/reference/MSO_AUTO_SHAPE_TYPE.md)
  : MSO_AUTO_SHAPE_TYPE — preset geometry names for autoshapes

- [`MSO_SHAPE_TYPE`](https://markheckmann.github.io/rpptx/reference/MSO_SHAPE_TYPE.md)
  : MSO_SHAPE_TYPE — shape type classification constants

- [`MSO_CONNECTOR_TYPE`](https://markheckmann.github.io/rpptx/reference/MSO_CONNECTOR_TYPE.md)
  : Connector type constants (MSO_CONNECTOR_TYPE)

- [`PP_PLACEHOLDER`](https://markheckmann.github.io/rpptx/reference/PP_PLACEHOLDER.md)
  : PP_PLACEHOLDER — placeholder type string constants

- [`MSO_THEME_COLOR`](https://markheckmann.github.io/rpptx/reference/MSO_THEME_COLOR.md)
  : MSO_THEME_COLOR — theme color XML value strings

- [`MSO_FILL`](https://markheckmann.github.io/rpptx/reference/MSO_FILL.md)
  : MSO_FILL — fill type strings

- [`MSO_LINE_DASH_STYLE`](https://markheckmann.github.io/rpptx/reference/MSO_LINE_DASH_STYLE.md)
  :

  MSO_LINE_DASH_STYLE — preset dash pattern strings for
  `<a:prstDash val="..."/>`

- [`PP_PARAGRAPH_ALIGNMENT`](https://markheckmann.github.io/rpptx/reference/PP_PARAGRAPH_ALIGNMENT.md)
  : PP_PARAGRAPH_ALIGNMENT — paragraph alignment

- [`MSO_AUTO_SIZE`](https://markheckmann.github.io/rpptx/reference/MSO_AUTO_SIZE.md)
  : MSO_AUTO_SIZE — text-box auto-sizing behaviour

- [`MSO_VERTICAL_ANCHOR`](https://markheckmann.github.io/rpptx/reference/MSO_VERTICAL_ANCHOR.md)
  : MSO_VERTICAL_ANCHOR — vertical text anchor

- [`XL_CHART_TYPE`](https://markheckmann.github.io/rpptx/reference/XL_CHART_TYPE.md)
  : XL_CHART_TYPE — chart type enumeration

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
- [`qn()`](https://markheckmann.github.io/rpptx/reference/qn.md) : Get
  the Clark-notation qualified tag name for a namespace-prefixed tag
