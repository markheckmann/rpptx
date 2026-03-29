# Create or open a PowerPoint presentation

Create or open a PowerPoint presentation

## Usage

``` r
pptx_presentation(pptx = NULL)
```

## Arguments

- pptx:

  Path to a .pptx file, or NULL to use the default template.

## Value

A Presentation object.

## Examples

``` r
# A comprehensive end-to-end example that demonstrates the rpptx API:
# creating a presentation from the built-in template, building five slides
# covering title placeholders, rich text formatting, autoshapes, connectors,
# slide backgrounds, tables, charts (column, line, pie, XY), and saving.

prs <- pptx_presentation()

# ── Presentation dimensions ───────────────────────────────────────────────────
# Default template is widescreen (13.33" × 7.5")
cat("Slide size:", as_inches(prs$slide_width), "×",
    as_inches(prs$slide_height), "inches\n")
#> Slide size: 10 × 7.5 inches

# ── Available slide layouts ───────────────────────────────────────────────────
layouts  <- prs$slide_layouts
n_layouts <- length(layouts)
cat("Layouts available:", n_layouts, "\n")
#> Layouts available: 11 
for (i in seq_len(n_layouts)) cat("  ", i, ":", layouts[[i]]$name, "\n")
#>    1 : Title Slide 
#>    2 : Title and Content 
#>    3 : Section Header 
#>    4 : Two Content 
#>    5 : Comparison 
#>    6 : Title Only 
#>    7 : Blank 
#>    8 : Content with Caption 
#>    9 : Picture with Caption 
#>    10 : Title and Vertical Text 
#>    11 : Vertical Title and Text 

# ── Helpers ───────────────────────────────────────────────────────────────────
blank_layout <- layouts[[6]]   # "Blank" layout — no placeholder clutter

# ── Slide 1: Title slide ──────────────────────────────────────────────────────
title_layout <- layouts[[1]]
slide1 <- prs$slides$add_slide(title_layout)

ph <- slide1$placeholders
ph[[1]]$text_frame$text <- "rpptx Showcase"
ph[[2]]$text_frame$text <- "An R port of python-pptx"

# Style the title placeholder text
title_para <- ph[[1]]$text_frame$paragraphs[[1]]
title_para$alignment  <- PP_PARAGRAPH_ALIGNMENT$CENTER
title_para$font$bold  <- TRUE
title_para$font$size  <- Pt(40)
title_para$font$color$rgb <- RGBColor(0x1F, 0x49, 0x7D)

# Style the subtitle
sub_para <- ph[[2]]$text_frame$paragraphs[[1]]
sub_para$alignment <- PP_PARAGRAPH_ALIGNMENT$CENTER
sub_para$font$size  <- Pt(24)
sub_para$font$color$rgb <- RGBColor(0x40, 0x40, 0x40)

# ── Slide 2: Rich text & autoshapes ───────────────────────────────────────────
slide2 <- prs$slides$add_slide(blank_layout)

# Coloured slide background
slide2$background$fill$solid()
slide2$background$fill$fore_color$rgb <- RGBColor(0xF2, 0xF7, 0xFF)

# --- Textbox with multi-paragraph, multi-run content ---
txb <- slide2$shapes$add_textbox(
  Inches(0.4), Inches(0.3), Inches(6.0), Inches(1.6)
)
tf <- txb$text_frame
tf$word_wrap <- TRUE

# First paragraph: large heading
p1 <- tf$paragraphs[[1]]
p1$alignment <- PP_PARAGRAPH_ALIGNMENT$LEFT
r1 <- p1$add_run()
r1$text <- "Text formatting demo"
r1$font$bold  <- TRUE
r1$font$size  <- Pt(22)
r1$font$color$rgb <- RGBColor(0x1F, 0x49, 0x7D)


# Second paragraph: mixed runs
p2 <- tf$add_paragraph()
p2$space_before <- Pt(6)
r2a <- p2$add_run(); r2a$text <- "Normal  "
r2b <- p2$add_run(); r2b$text <- "bold  ";  r2b$font$bold   <- TRUE
r2c <- p2$add_run(); r2c$text <- "italic  "; r2c$font$italic <- TRUE
r2d <- p2$add_run(); r2d$text <- "underline"
r2d$font$underline <- TRUE
r2d$font$color$rgb <- RGBColor(0x70, 0x30, 0xA0)

# Third paragraph: different font
p3 <- tf$add_paragraph()
p3$space_before <- Pt(4)
r3 <- p3$add_run()
r3$text <- "Monospaced: Courier New, 14 pt"
r3$font$name <- "Courier New"
r3$font$size <- Pt(14)
r3$font$color$rgb <- RGBColor(0xC0, 0x40, 0x00)

# --- Rounded rectangle with gradient fill ---
rrect <- slide2$shapes$add_shape(
  MSO_AUTO_SHAPE_TYPE$ROUNDED_RECTANGLE,
  Inches(6.8), Inches(0.3), Inches(2.8), Inches(1.6)
)
rrect$name <- "GradientRoundRect"
rrect$fill$gradient()
stops <- rrect$fill$gradient_stops
stops[[1]]$position          <- 0.0
stops[[1]]$color$rgb         <- RGBColor(0x1F, 0x49, 0x7D)
stops[[2]]$position          <- 1.0
stops[[2]]$color$rgb         <- RGBColor(0x70, 0xB0, 0xFF)
rrect$fill$gradient_angle    <- 135
rrect$line$width             <- Pt(0)   # no border

inner_tf  <- rrect$text_frame
inner_tf$vertical_anchor <- MSO_VERTICAL_ANCHOR$MIDDLE
inner_p   <- inner_tf$paragraphs[[1]]
inner_p$alignment <- PP_PARAGRAPH_ALIGNMENT$CENTER
inner_r   <- inner_p$add_run()
inner_r$text       <- "Gradient fill"
inner_r$font$bold  <- TRUE
inner_r$font$size  <- Pt(16)
inner_r$font$color$rgb <- RGBColor(0xFF, 0xFF, 0xFF)

# --- Oval with solid fill ---
oval <- slide2$shapes$add_shape(
  MSO_AUTO_SHAPE_TYPE$OVAL,
  Inches(0.4), Inches(2.2), Inches(2.0), Inches(2.0)
)
oval$fill$solid()
oval$fill$fore_color$rgb <- RGBColor(0xFF, 0xC0, 0x00)
oval$line$width          <- Pt(2)
oval$line$color$rgb      <- RGBColor(0x80, 0x60, 0x00)
oval_r <- oval$text_frame$paragraphs[[1]]$add_run()
oval_r$text           <- "Oval"
oval_r$font$bold      <- TRUE
oval_r$font$size      <- Pt(18)
oval_r$font$color$rgb <- RGBColor(0x40, 0x30, 0x00)
oval$text_frame$paragraphs[[1]]$alignment <- PP_PARAGRAPH_ALIGNMENT$CENTER

# --- Pentagon arrow pointing right ---
arrow <- slide2$shapes$add_shape(
  MSO_AUTO_SHAPE_TYPE$PENTAGON,
  Inches(2.8), Inches(2.4), Inches(2.4), Inches(1.5)
)
arrow$fill$solid()
arrow$fill$fore_color$rgb <- RGBColor(0x70, 0xB0, 0x50)
arrow$line$color$rgb      <- RGBColor(0x30, 0x60, 0x20)
arrow$line$width          <- Pt(1.5)
arrow_r <- arrow$text_frame$paragraphs[[1]]$add_run()
arrow_r$text      <- "Pentagon"
arrow_r$font$size <- Pt(14)
arrow$text_frame$paragraphs[[1]]$alignment <- PP_PARAGRAPH_ALIGNMENT$CENTER

# --- Connector linking the two shapes ---
conn <- slide2$shapes$add_connector(
  MSO_CONNECTOR_TYPE$STRAIGHT,
  Inches(2.4),  Inches(3.2),
  Inches(2.8),  Inches(3.15)
)
conn$line$width          <- Pt(2.0)
conn$line$color$rgb      <- RGBColor(0x20, 0x20, 0x20)
conn$line$dash_style     <- MSO_LINE_DASH_STYLE$DASH

# --- Star shape ---
star <- slide2$shapes$add_shape(
  MSO_AUTO_SHAPE_TYPE$STAR_5_POINT,
  Inches(5.4), Inches(2.2), Inches(2.0), Inches(2.0)
)
star$fill$solid()
star$fill$fore_color$rgb <- RGBColor(0xFF, 0x40, 0x40)
star$line$width          <- Pt(1)
star$line$color$rgb      <- RGBColor(0x80, 0x00, 0x00)
star$rotation            <- 15   # degrees clockwise

# --- Patterned fill shape ---
pat_box <- slide2$shapes$add_shape(
  MSO_AUTO_SHAPE_TYPE$RECTANGLE,
  Inches(7.8), Inches(2.3), Inches(1.8), Inches(1.2)
)
pat_box$fill$patterned()
pat_box$fill$pattern      <- MSO_PATTERN_TYPE$DIAGONAL_BRICK
pat_box$fill$fore_color$rgb <- RGBColor(0x1F, 0x49, 0x7D)
pat_box$fill$back_color$rgb <- RGBColor(0xCC, 0xDD, 0xFF)
pat_box$line$width          <- Pt(1)


# ── Slide 3: Table ────────────────────────────────────────────────────────────
slide3 <- prs$slides$add_slide(blank_layout)

# Title textbox
title3 <- slide3$shapes$add_textbox(
  Inches(0.4), Inches(0.2), Inches(9.2), Inches(0.7)
)
tr3 <- title3$text_frame$paragraphs[[1]]$add_run()
tr3$text      <- "Feature matrix"
tr3$font$bold <- TRUE
tr3$font$size <- Pt(26)
tr3$font$color$rgb <- RGBColor(0x1F, 0x49, 0x7D)

# Add 5 × 4 table
gf_tbl <- slide3$shapes$add_table(
  5L, 4L,
  Inches(0.4), Inches(1.0), Inches(9.2), Inches(3.6)
)
tbl <- gf_tbl$table

# Column widths
tbl$columns[[1]]$width <- Inches(3.2)
tbl$columns[[2]]$width <- Inches(2.0)
tbl$columns[[3]]$width <- Inches(2.0)
tbl$columns[[4]]$width <- Inches(2.0)

# Row heights
for (i in seq_len(5)) tbl$rows[[i]]$height <- Inches(0.7)

# Header row — dark blue background, white bold text
headers <- c("Feature", "Status", "Since", "Notes")
for (j in seq_along(headers)) {
  hcell <- tbl$cell(1L, j)
  hcell$text <- headers[j]
  hcell$fill$solid()
  hcell$fill$fore_color$rgb <- RGBColor(0x1F, 0x49, 0x7D)
  hp <- hcell$text_frame$paragraphs[[1]]
  hp$alignment        <- PP_PARAGRAPH_ALIGNMENT$CENTER
  hp$font$bold        <- TRUE
  hp$font$color$rgb   <- RGBColor(0xFF, 0xFF, 0xFF)
  hp$font$size        <- Pt(13)
}

# Data rows
rows_data <- list(
  c("Text & fonts",     "complete",  "0.1.0", "Bold, italic, colour, size"),
  c("Shapes & fills",   "complete",  "0.1.0", "Solid, gradient, pattern"),
  c("Tables",           "complete",  "0.1.0", "Cell merge, banding, style"),
  c("Charts",           "complete",  "0.1.0", "Column, bar, line, pie, XY")
)
for (i in seq_along(rows_data)) {
  row_idx <- i + 1L
  fill_row <- (i %% 2L == 0L)
  for (j in seq_along(rows_data[[i]])) {
    dcell <- tbl$cell(row_idx, j)
    dcell$text <- rows_data[[i]][j]
    dp <- dcell$text_frame$paragraphs[[1]]
    dp$font$size <- Pt(12)
    if (fill_row) {
      dcell$fill$solid()
      dcell$fill$fore_color$rgb <- RGBColor(0xE8, 0xF0, 0xFE)
    }
  }
  # Colour "complete" green
  status_cell <- tbl$cell(row_idx, 2L)
  status_cell$text_frame$paragraphs[[1]]$font$color$rgb <-
    RGBColor(0x18, 0x80, 0x38)
  status_cell$text_frame$paragraphs[[1]]$font$bold <- TRUE
}

# Enable alternating row banding flag
tbl$horz_banding <- FALSE   # we applied manual banding above


# ── Slide 4: Clustered column chart + line chart ──────────────────────────────
slide4 <- prs$slides$add_slide(blank_layout)

# Title
title4 <- slide4$shapes$add_textbox(
  Inches(0.4), Inches(0.15), Inches(9.2), Inches(0.6)
)
t4r <- title4$text_frame$paragraphs[[1]]$add_run()
t4r$text <- "Quarterly sales — column & line charts"
t4r$font$bold <- TRUE; t4r$font$size <- Pt(22)
t4r$font$color$rgb <- RGBColor(0x1F, 0x49, 0x7D)

# --- Clustered column chart (left half) ---
col_data <- CategoryChartData$new()
col_data$categories <- c("Q1", "Q2", "Q3", "Q4")
col_data$add_series("Product A", c(120, 145, 180, 210))
col_data$add_series("Product B", c( 85,  95, 130, 160))
col_data$add_series("Product C", c( 50,  75,  90, 115))

gf_col <- slide4$shapes$add_chart(
  XL_CHART_TYPE$COLUMN_CLUSTERED,
  Inches(0.3), Inches(0.85), Inches(6.0), Inches(5.8),
  col_data
)
chart_col <- gf_col$chart
chart_col$has_title  <- TRUE
chart_col$has_legend <- TRUE
chart_col$chart_title$text_frame$text <- "Unit sales"

va <- chart_col$value_axis
va$minimum_scale         <- 0
va$maximum_scale         <- 250
va$has_major_gridlines   <- TRUE
va$minor_tick_mark       <- XL_TICK_MARK$NONE
va$major_tick_mark       <- XL_TICK_MARK$OUTSIDE

ca <- chart_col$category_axis
ca$tick_labels$font$size <- Pt(10)

# --- Line chart (right half, same data) ---
line_data <- CategoryChartData$new()
line_data$categories <- c("Q1", "Q2", "Q3", "Q4")
line_data$add_series("Product A", c(120, 145, 180, 210))
line_data$add_series("Product B", c( 85,  95, 130, 160))

gf_line <- slide4$shapes$add_chart(
  XL_CHART_TYPE$LINE_MARKERS,
  Inches(6.5), Inches(0.85), Inches(3.2), Inches(2.7),
  line_data
)
chart_line <- gf_line$chart
chart_line$has_title  <- TRUE
chart_line$has_legend <- FALSE
chart_line$chart_title$text_frame$text <- "Trend"


# ── Slide 5: Pie chart + XY scatter ──────────────────────────────────────────
slide5 <- prs$slides$add_slide(blank_layout)

title5 <- slide5$shapes$add_textbox(
  Inches(0.4), Inches(0.15), Inches(9.2), Inches(0.6)
)
t5r <- title5$text_frame$paragraphs[[1]]$add_run()
t5r$text <- "Pie chart and XY scatter"
t5r$font$bold <- TRUE; t5r$font$size <- Pt(22)
t5r$font$color$rgb <- RGBColor(0x1F, 0x49, 0x7D)

# --- Pie chart ---
pie_data <- CategoryChartData$new()
pie_data$categories <- c("APAC", "EMEA", "Americas", "Other")
pie_data$add_series("Revenue share", c(38, 31, 25, 6))

gf_pie <- slide5$shapes$add_chart(
  XL_CHART_TYPE$PIE,
  Inches(0.3), Inches(0.85), Inches(5.5), Inches(5.8),
  pie_data
)
chart_pie <- gf_pie$chart
chart_pie$has_title  <- TRUE
chart_pie$has_legend <- TRUE
chart_pie$chart_title$text_frame$text <- "Revenue by region (%)"

# Show percentage data labels
plot_pie <- chart_pie$plots[[1]]
plot_pie$has_data_labels  <- TRUE
dl <- plot_pie$data_labels
dl$number_format          <- "0%"
dl$show_percentage        <- TRUE
dl$show_value             <- FALSE
dl$show_series_name       <- FALSE

# --- XY scatter chart ---
xy_data <- XyChartData$new()
series_xy <- xy_data$add_series("Observations")
xy_points <- data.frame(
  x = c(1.0, 1.5, 2.2, 2.8, 3.5, 4.0, 4.6, 5.1, 5.9, 6.4),
  y = c(2.1, 2.9, 4.0, 4.8, 6.3, 7.0, 8.2, 8.9, 10.4, 11.1)
)
for (k in seq_len(nrow(xy_points))) {
  series_xy$add_data_point(xy_points$x[k], xy_points$y[k])
}

gf_xy <- slide5$shapes$add_chart(
  XL_CHART_TYPE$XY_SCATTER_LINES_NO_MARKERS,
  Inches(6.0), Inches(0.85), Inches(3.7), Inches(5.8),
  xy_data
)
chart_xy <- gf_xy$chart
chart_xy$has_title  <- TRUE
chart_xy$has_legend <- FALSE
chart_xy$chart_title$text_frame$text <- "X–Y scatter"
chart_xy$value_axis$has_major_gridlines    <- TRUE
chart_xy$category_axis$has_major_gridlines <- TRUE

# ── Save ──────────────────────────────────────────────────────────────────────
out_path <- tempfile(fileext = ".pptx")
prs$save(out_path)
cat("Saved presentation to:", out_path, "\n")
#> Saved presentation to: /tmp/RtmplAuXca/file1e456610248c.pptx 
cat("Slides:", length(prs$slides), "\n")
#> Slides: 5 
```
