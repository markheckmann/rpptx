# rpptx 0.1.0

## New features

* Full read/write support for PowerPoint (.pptx) files.

### Presentation
* `pptx_presentation()` — open or create a presentation
* `Presentation$save()` — save to disk
* `Presentation$slide_width`, `$slide_height` — slide dimensions
* `Presentation$slides`, `$slide_layouts`, `$slide_masters` — collections
* `Presentation$core_properties` — document metadata

### Slides
* `Slides$add_slide()` — add a new slide from a layout
* `Slide$shapes` — shape collection for a slide
* `Slide$placeholders` — placeholder collection
* `SlideLayout$placeholders`, `SlideMaster$placeholders` — layout/master placeholders
* Placeholder inheritance: slides inherit from layout from master

### Shapes — read
* All shape types supported: `Shape`, `Picture`, `GraphicFrame`, `GroupShape`,
  `Connector`, `SlidePlaceholder`, `LayoutPlaceholder`, `MasterPlaceholder`
* Position and size: `$left`, `$top`, `$width`, `$height`
* `$name`, `$shape_id`, `$shape_type`, `$has_text_frame`

### Shapes — create
* `slide$shapes$add_shape()` — autoshape
* `slide$shapes$add_textbox()` — text box
* `slide$shapes$add_picture()` — image (PNG, JPEG, etc. via `magick`)
* `slide$shapes$add_table()` — table
* `slide$shapes$add_chart()` — chart
* `slide$shapes$add_connector()` — connector
* `slide$shapes$add_group_shape()` — group existing shapes
* `slide$shapes$build_freeform()` — custom-geometry freeform shapes

### Text
* `shape$text_frame` — `TextFrame` with `$paragraphs`, `$text`, `$word_wrap`,
  `$vertical_anchor`, margins
* `paragraph$runs`, `paragraph$add_run()`, `paragraph$alignment`
* `run$text`, `run$font` — `Font` with `$bold`, `$italic`, `$underline`,
  `$size`, `$name`, `$color`

### Tables
* `shape$table` — `Table` with `$rows`, `$columns`, `$cell(row, col)`
* `cell$text`, `cell$text_frame`, `cell$fill`
* `as.data.frame(table)` — extract cell text as a data frame

### Charts
* Supported chart types: column, bar, line, pie, area, scatter (XY), bubble,
  radar, doughnut, and their variants
* `CategoryChartData`, `XyChartData`, `BubbleChartData` — data containers
* Chart proxy: `$chart_title`, `$plots`, `$series`, `$legend`

### Formatting (DML)
* `shape$fill` — `FillFormat` with solid, gradient, pattern, blip, no-fill
* `shape$line` — `LineFormat` with width, dash style
* `ColorFormat`, `RGBColor` — color specification
* `MSO_THEME_COLOR`, `MSO_FILL`, `MSO_LINE_DASH_STYLE` enumerations

### Enumerations
* `MSO_AUTO_SHAPE_TYPE`, `MSO_SHAPE_TYPE`, `PP_PLACEHOLDER`
* `MSO_CONNECTOR_TYPE`, `MSO_THEME_COLOR`
* `PP_PARAGRAPH_ALIGNMENT` / `PP_ALIGN`, `MSO_AUTO_SIZE`, `MSO_VERTICAL_ANCHOR`
* `XL_CHART_TYPE` (73 chart types)

### Units
* `Inches()`, `Cm()`, `Pt()`, `Mm()`, `Emu()`, `Centipoints()` — length unit constructors

### Display
* `print()` / `format()` methods for all major objects: `Presentation`, `Slide`,
  `SlideLayout`, `SlideMaster`, shapes, `Table`, `TextFrame`
