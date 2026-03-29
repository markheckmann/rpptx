# rpptx

**rpptx** is an R package for creating, reading, and updating PowerPoint
(.pptx) presentations. It is a close R port of the Python
[python-pptx](https://python-pptx.readthedocs.io/) library, providing an
R6-based object model that mirrors the Python API.

## Installation

```r
# Development version from GitHub:
# install.packages("pak")
pak::pak("markheckmann/rpptx")
```

## Quick start

```r
library(rpptx)

# ---- Open / create -----------------------------------------------------------
prs <- pptx_presentation()                    # blank presentation
prs <- pptx_presentation("my_file.pptx")      # open existing

# ---- Add a slide -------------------------------------------------------------
layout <- prs$slide_layouts[[1]]              # "Title Slide" layout
slide  <- prs$slides$add_slide(layout)

# ---- Add shapes --------------------------------------------------------------
# Text box
box <- slide$shapes$add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
tf  <- box$text_frame
tf$text <- "Hello, rpptx!"

# Autoshape (filled rectangle)
shp <- slide$shapes$add_shape(
  MSO_AUTO_SHAPE_TYPE$ROUNDED_RECTANGLE,
  Inches(1), Inches(2.5), Inches(3), Inches(1.5)
)
shp$fill$solid()
shp$fill$fore_color$rgb <- RGBColor(0x4F, 0x81, 0xBD)

# ---- Save --------------------------------------------------------------------
prs$save("output.pptx")
```

## Features

### Presentations
```r
prs$slide_width   # Emu (use .emu_in() or as.numeric() / 914400)
prs$slide_height
prs$slides        # collection: prs$slides[[1]], length(prs$slides)
prs$slide_layouts # 11 default layouts
prs$slide_masters
prs$core_properties$title <- "My Presentation"
```

### Text and fonts
```r
tf <- shape$text_frame
tf$text <- "Replace all text"
tf$word_wrap <- TRUE
tf$vertical_anchor <- MSO_VERTICAL_ANCHOR$MIDDLE

para <- tf$paragraphs[[1]]
para$alignment <- PP_ALIGN$CENTER
run  <- para$add_run()
run$text <- "Bold blue text"
run$font$bold  <- TRUE
run$font$size  <- Pt(24)
run$font$color$rgb <- RGBColor(0xFF, 0x00, 0x00)
```

### Tables
```r
gf  <- slide$shapes$add_table(3, 4, Inches(1), Inches(2), Inches(8), Inches(3))
tbl <- gf$table

# Set cell text
cell <- tbl$cell(1, 1)
cell$text <- "Header"

# Style a cell
cell$fill$solid()
cell$fill$fore_color$rgb <- RGBColor(0x4F, 0x81, 0xBD)

# Convert to data frame
df <- as.data.frame(tbl)
```

### Charts
```r
# Column chart
chart_data <- CategoryChartData$new()
chart_data$categories <- c("East", "West", "North")
chart_data$add_series("Q1", c(19.2, 21.4, 16.7))
chart_data$add_series("Q2", c(22.3, 28.6, 15.2))

gf <- slide$shapes$add_chart(
  XL_CHART_TYPE$COLUMN_CLUSTERED,
  Inches(1), Inches(1), Inches(8), Inches(5),
  chart_data
)

# Scatter (XY) chart
xy_data <- XyChartData$new()
s <- xy_data$add_series("Series 1")
s$add_data_point(1.0, 2.3)
s$add_data_point(1.5, 1.8)
slide$shapes$add_chart(XL_CHART_TYPE$XY_SCATTER, ...)

# Bubble chart
bd <- BubbleChartData$new()
s  <- bd$add_series("Bubbles")
s$add_data_point(x = 1.0, y = 2.0, size = 10)
s$add_data_point(x = 2.0, y = 3.0, size = 20)
slide$shapes$add_chart(XL_CHART_TYPE$BUBBLE, ...)
```

### Images
```r
pic <- slide$shapes$add_picture(
  "photo.jpg",
  left  = Inches(1),
  top   = Inches(1),
  width = Inches(4)     # height auto-calculated if omitted
)
```

### Freeform shapes
```r
# Local coordinate system where 100 units = 1 inch
scale <- Inches(1) / 100

ff  <- slide$shapes$build_freeform(start_x = 0, start_y = 0, scale = scale)
ff$add_line_segments(list(c(100, 0), c(50, 87), c(0, 0)))   # triangle
shp <- ff$convert_to_shape(origin_x = Inches(1), origin_y = Inches(1))
```

### Fill and line formatting
```r
# Solid fill
shp$fill$solid()
shp$fill$fore_color$rgb <- RGBColor(0xFF, 0x00, 0x00)

# Theme color
shp$fill$fore_color$theme_color <- MSO_THEME_COLOR$ACCENT_1

# Line
shp$line$width <- Pt(2)
shp$line$dash_style <- MSO_LINE_DASH_STYLE$DASH

# No fill
shp$fill$background()
```

### Connectors
```r
conn <- slide$shapes$add_connector(
  MSO_CONNECTOR_TYPE$STRAIGHT,
  begin_x = Inches(1), begin_y = Inches(1),
  end_x   = Inches(3), end_y   = Inches(3)
)
```

### Group shapes
```r
# Group two existing shapes
grp <- slide$shapes$add_group_shape(list(shape1, shape2))
```

### Print methods
```r
prs               # <Presentation>  slides: 3  size: 10.00 × 7.50 in ...
slide             # <Slide>  9 shapes  ...
shape             # <AutoShape>  "Rectangle 1"  3.00×2.00 in @ 1.00, 1.00 in
tbl               # <Table>  3 rows × 4 cols
```

## Unit helpers

```r
Inches(1.5)       # 1371600 EMU
Cm(2.54)          # 914400 EMU (= 1 inch)
Pt(18)            # 228600 EMU
Mm(25.4)          # 914400 EMU
Emu(914400)       # explicit EMU
```

## Relationship to python-pptx

rpptx is a direct R port of [python-pptx](https://python-pptx.readthedocs.io/).
The object model, property names, and API are intentionally kept as close as
possible to the Python original. Key differences:

| python-pptx | rpptx |
|---|---|
| `Presentation()` | `pptx_presentation()` |
| `prs.slides[0]` | `prs$slides[[1]]` (1-based) |
| `shape.text_frame` | `shape$text_frame` |
| `font.bold = True` | `font$bold <- TRUE` |
| `Inches(1.5)` | `Inches(1.5)` |
| `MSO_SHAPE.RECTANGLE` | `MSO_AUTO_SHAPE_TYPE$RECTANGLE` |

## License

MIT
