# Tests for Phase 8: Charts
# Covers: enum-chart, chart-data, chart-xlsx, chart-xmlwriter, parts-chart,
#         add_chart integration via SlideShapes.


# ============================================================================
# XL_CHART_TYPE enum
# ============================================================================

describe("XL_CHART_TYPE", {
  it("contains expected integer constants", {
    expect_equal(XL_CHART_TYPE$COLUMN_CLUSTERED, 51L)
    expect_equal(XL_CHART_TYPE$BAR_CLUSTERED,    57L)
    expect_equal(XL_CHART_TYPE$LINE,              4L)
    expect_equal(XL_CHART_TYPE$PIE,               5L)
    expect_equal(XL_CHART_TYPE$DOUGHNUT,      -4120L)
    expect_equal(XL_CHART_TYPE$XY_SCATTER,    -4169L)
    expect_equal(XL_CHART_TYPE$BUBBLE,           15L)
    expect_equal(XL_CHART_TYPE$RADAR,         -4151L)
  })
})


# ============================================================================
# .excel_col_letter helper
# ============================================================================

describe(".excel_col_letter", {
  it("converts 1-based column number to letter", {
    expect_equal(.excel_col_letter(1L),  "A")
    expect_equal(.excel_col_letter(26L), "Z")
    expect_equal(.excel_col_letter(27L), "AA")
    expect_equal(.excel_col_letter(52L), "AZ")
    expect_equal(.excel_col_letter(53L), "BA")
    expect_equal(.excel_col_letter(703L), "AAA")
  })

  it("errors on out-of-range input", {
    expect_error(.excel_col_letter(0L))
    expect_error(.excel_col_letter(16385L))
  })
})


# ============================================================================
# CategoryChartData / Categories / Category
# ============================================================================

describe("CategoryChartData", {
  it("accepts category labels and series values", {
    cd <- CategoryChartData$new()
    cd$categories <- c("Q1", "Q2", "Q3")
    cd$add_series("Sales", c(100, 200, 150))

    cats <- cd$categories
    expect_equal(length(cats), 3L)
    expect_equal(cats$get(1L)$label, "Q1")
    expect_equal(cats$get(3L)$label, "Q3")
    expect_equal(length(cd), 1L)

    series <- cd$get(1L)
    expect_equal(series$name, "Sales")
    expect_equal(series$values, c(100, 200, 150))
    expect_equal(series$index, 0L)
    expect_equal(series$n_points, 3L)
  })

  it("supports multiple series", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    cd$add_series("S1", c(1, 2))
    cd$add_series("S2", c(3, 4))

    expect_equal(length(cd), 2L)
    expect_equal(cd$get(2L)$index, 1L)
  })

  it("ChartData is an alias for CategoryChartData", {
    cd <- ChartData$new()
    expect_true(inherits(cd, "CategoryChartData"))
  })
})


describe("CategoryWorkbookWriter cell references", {
  it("categories_ref covers depth=1 categories", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B", "C")
    cd$add_series("S", c(1, 2, 3))
    w <- CategoryWorkbookWriter$new(cd)
    expect_equal(w$categories_ref(), "Sheet1!$A$2:$A$4")
  })

  it("series_name_ref points to header row", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    s <- cd$add_series("Revenue", c(10, 20))
    w <- CategoryWorkbookWriter$new(cd)
    expect_equal(w$series_name_ref(s), "Sheet1!$B$1")
  })

  it("values_ref covers data rows", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    s <- cd$add_series("Revenue", c(10, 20))
    w <- CategoryWorkbookWriter$new(cd)
    expect_equal(w$values_ref(s), "Sheet1!$B$2:$B$3")
  })

  it("second series is in the next column", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    cd$add_series("S1", c(1, 2))
    s2 <- cd$add_series("S2", c(3, 4))
    w <- CategoryWorkbookWriter$new(cd)
    expect_equal(w$series_name_ref(s2), "Sheet1!$C$1")
    expect_equal(w$values_ref(s2), "Sheet1!$C$2:$C$3")
  })
})


describe("XyWorkbookWriter cell references", {
  it("series_name_ref points to B1 for first series", {
    cd <- XyChartData$new()
    s  <- cd$add_series("Series 1")
    s$add_data_point(1, 2)
    s$add_data_point(3, 4)
    w  <- XyWorkbookWriter$new(cd)
    expect_equal(w$series_name_ref(s), "Sheet1!$B$1")
  })

  it("x_values_ref and y_values_ref match data rows", {
    cd <- XyChartData$new()
    s  <- cd$add_series("S1")
    s$add_data_point(1, 2)
    s$add_data_point(3, 4)
    w  <- XyWorkbookWriter$new(cd)
    expect_equal(w$x_values_ref(s), "Sheet1!$A$2:$A$3")
    expect_equal(w$y_values_ref(s), "Sheet1!$B$2:$B$3")
  })
})


describe("BubbleWorkbookWriter cell references", {
  it("bubble_sizes_ref points to column C", {
    cd <- BubbleChartData$new()
    s  <- cd$add_series("Bubbles")
    s$add_data_point(1, 2, 10)
    s$add_data_point(3, 4, 20)
    w  <- BubbleWorkbookWriter$new(cd)
    expect_equal(w$bubble_sizes_ref(s), "Sheet1!$C$2:$C$3")
  })
})


# ============================================================================
# chart_xml_writer — factory and XML structure
# ============================================================================

describe("chart_xml_writer", {
  it("errors for an unsupported chart type", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A")
    cd$add_series("S", c(1))
    expect_error(chart_xml_writer(999L, cd))
  })

  it("generates XML starting with <?xml for column clustered", {
    cd <- CategoryChartData$new()
    cd$categories <- c("East", "West", "North")
    cd$add_series("Sales", c(100, 200, 150))
    xml <- chart_xml_writer(XL_CHART_TYPE$COLUMN_CLUSTERED, cd)$xml
    expect_true(startsWith(xml, "<?xml"))
    expect_true(grepl("c:barChart", xml, fixed = TRUE))
    expect_true(grepl('barDir val="col"', xml, fixed = TRUE))
    expect_true(grepl('grouping val="clustered"', xml, fixed = TRUE))
    expect_true(grepl("East", xml, fixed = TRUE))
    expect_true(grepl("Sales", xml, fixed = TRUE))
  })

  it("generates bar chart XML for BAR_CLUSTERED", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    cd$add_series("S", c(1, 2))
    xml <- chart_xml_writer(XL_CHART_TYPE$BAR_CLUSTERED, cd)$xml
    expect_true(grepl('barDir val="bar"', xml, fixed = TRUE))
  })

  it("generates stacked bar XML for COLUMN_STACKED", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A")
    cd$add_series("S", c(1))
    xml <- chart_xml_writer(XL_CHART_TYPE$COLUMN_STACKED, cd)$xml
    expect_true(grepl('grouping val="stacked"', xml, fixed = TRUE))
    expect_true(grepl('overlap val="100"', xml, fixed = TRUE))
  })

  it("generates area chart XML for AREA", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    cd$add_series("S", c(1, 2))
    xml <- chart_xml_writer(XL_CHART_TYPE$AREA, cd)$xml
    expect_true(grepl("c:areaChart", xml, fixed = TRUE))
    expect_true(grepl('grouping val="standard"', xml, fixed = TRUE))
  })

  it("generates line chart XML for LINE", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    cd$add_series("S", c(1, 2))
    xml <- chart_xml_writer(XL_CHART_TYPE$LINE, cd)$xml
    expect_true(grepl("c:lineChart", xml, fixed = TRUE))
    expect_true(grepl('symbol val="none"', xml, fixed = TRUE))
  })

  it("generates line chart with markers for LINE_MARKERS", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A")
    cd$add_series("S", c(1))
    xml <- chart_xml_writer(XL_CHART_TYPE$LINE_MARKERS, cd)$xml
    expect_false(grepl('symbol val="none"', xml, fixed = TRUE))
  })

  it("generates pie chart XML for PIE", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B", "C")
    cd$add_series("S", c(10, 20, 30))
    xml <- chart_xml_writer(XL_CHART_TYPE$PIE, cd)$xml
    expect_true(grepl("c:pieChart", xml, fixed = TRUE))
    expect_false(grepl("explosion", xml, fixed = TRUE))
  })

  it("generates exploded pie XML for PIE_EXPLODED", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    cd$add_series("S", c(1, 2))
    xml <- chart_xml_writer(XL_CHART_TYPE$PIE_EXPLODED, cd)$xml
    expect_true(grepl('explosion val="25"', xml, fixed = TRUE))
  })

  it("generates doughnut chart XML", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    cd$add_series("S", c(1, 2))
    xml <- chart_xml_writer(XL_CHART_TYPE$DOUGHNUT, cd)$xml
    expect_true(grepl("c:doughnutChart", xml, fixed = TRUE))
  })

  it("generates radar chart XML", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B", "C")
    cd$add_series("S", c(1, 2, 3))
    xml <- chart_xml_writer(XL_CHART_TYPE$RADAR, cd)$xml
    expect_true(grepl("c:radarChart", xml, fixed = TRUE))
    expect_true(grepl('radarStyle val="marker"', xml, fixed = TRUE))
  })

  it("generates filled radar XML for RADAR_FILLED", {
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    cd$add_series("S", c(1, 2))
    xml <- chart_xml_writer(XL_CHART_TYPE$RADAR_FILLED, cd)$xml
    expect_true(grepl('radarStyle val="filled"', xml, fixed = TRUE))
  })

  it("generates XY scatter XML", {
    cd <- XyChartData$new()
    s  <- cd$add_series("S")
    s$add_data_point(1, 2)
    s$add_data_point(3, 4)
    xml <- chart_xml_writer(XL_CHART_TYPE$XY_SCATTER, cd)$xml
    expect_true(grepl("c:scatterChart", xml, fixed = TRUE))
    expect_true(grepl("c:xVal", xml, fixed = TRUE))
    expect_true(grepl("c:yVal", xml, fixed = TRUE))
  })

  it("generates bubble chart XML", {
    cd <- BubbleChartData$new()
    s  <- cd$add_series("S")
    s$add_data_point(1, 2, 10)
    xml <- chart_xml_writer(XL_CHART_TYPE$BUBBLE, cd)$xml
    expect_true(grepl("c:bubbleChart", xml, fixed = TRUE))
    expect_true(grepl("c:bubbleSize", xml, fixed = TRUE))
  })
})


# ============================================================================
# CategoryChartData$xml_str and $xlsx_blob
# ============================================================================

describe("CategoryChartData xml_str and xlsx_blob", {
  it("xml_str returns valid XML string", {
    cd <- CategoryChartData$new()
    cd$categories <- c("East", "West")
    cd$add_series("Q1", c(100, 200))
    xml <- cd$xml_str(XL_CHART_TYPE$COLUMN_CLUSTERED)
    expect_true(nchar(xml) > 100L)
    expect_true(startsWith(xml, "<?xml"))
  })

  it("xlsx_blob returns raw bytes (requires openxlsx2)", {
    skip_if_not_installed("openxlsx2")
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B", "C")
    cd$add_series("Sales", c(10, 20, 30))
    blob <- cd$xlsx_blob
    expect_true(is.raw(blob))
    expect_gt(length(blob), 0L)
    # Verify it's a valid ZIP (xlsx): starts with PK signature
    expect_equal(blob[1:2], as.raw(c(0x50, 0x4b)))
  })
})


# ============================================================================
# XyChartData
# ============================================================================

describe("XyChartData", {
  it("adds series with data points", {
    cd <- XyChartData$new()
    s  <- cd$add_series("Series 1")
    s$add_data_point(1.0, 2.5)
    s$add_data_point(2.0, 3.5)
    expect_equal(s$n_points, 2L)
    expect_equal(s$x_values, c(1.0, 2.0))
    expect_equal(s$y_values, c(2.5, 3.5))
  })
})


# ============================================================================
# BubbleChartData
# ============================================================================

describe("BubbleChartData", {
  it("adds series with x, y, size data points", {
    cd <- BubbleChartData$new()
    s  <- cd$add_series("Bubbles")
    s$add_data_point(1, 2, 10)
    s$add_data_point(3, 4, 20)
    expect_equal(s$bubble_sizes, c(10, 20))
  })
})


# ============================================================================
# add_chart integration — full round trip
# ============================================================================

describe("SlideShapes$add_chart", {
  it("adds a chart graphicFrame to the slide", {
    skip_if_not_installed("openxlsx2")
    prs_path <- system.file("test_files", "no-slides.pptx", package = "rpptx")
    prs   <- pptx_presentation(prs_path)
    slide <- prs$slides$add_slide(prs$slide_masters[[1]]$slide_layouts[[1]])
    cd    <- CategoryChartData$new()
    cd$categories <- c("East", "West", "North")
    cd$add_series("Sales", c(100, 200, 150))

    gf <- slide$shapes$add_chart(
      XL_CHART_TYPE$COLUMN_CLUSTERED,
      Inches(1), Inches(1), Inches(6), Inches(4),
      cd
    )

    expect_true(inherits(gf, "GraphicFrame"))
    expect_true(isTRUE(gf$has_chart))
    expect_false(isTRUE(gf$has_table))
    expect_false(is.null(gf$chart_rId))
  })

  it("saves a presentation containing a chart without error", {
    skip_if_not_installed("openxlsx2")
    prs_path <- system.file("test_files", "no-slides.pptx", package = "rpptx")
    prs   <- pptx_presentation(prs_path)
    slide <- prs$slides$add_slide(prs$slide_masters[[1]]$slide_layouts[[1]])
    cd    <- CategoryChartData$new()
    cd$categories <- c("Q1", "Q2", "Q3", "Q4")
    cd$add_series("Revenue", c(100, 120, 140, 160))
    cd$add_series("Costs",   c(80,  90, 100, 110))

    slide$shapes$add_chart(
      XL_CHART_TYPE$COLUMN_CLUSTERED,
      Inches(1), Inches(1), Inches(8), Inches(5),
      cd
    )

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp), add = TRUE)
    expect_no_error(prs$save(tmp))
    expect_true(file.exists(tmp))
    expect_gt(file.size(tmp), 0L)
  })

  it("round-trips: chart rId survives save and reload", {
    skip_if_not_installed("openxlsx2")
    prs_path <- system.file("test_files", "no-slides.pptx", package = "rpptx")
    prs   <- pptx_presentation(prs_path)
    slide <- prs$slides$add_slide(prs$slide_masters[[1]]$slide_layouts[[1]])
    cd    <- CategoryChartData$new()
    cd$categories <- c("A", "B")
    cd$add_series("S1", c(1, 2))

    gf1   <- slide$shapes$add_chart(
      XL_CHART_TYPE$BAR_CLUSTERED,
      Inches(1), Inches(1), Inches(6), Inches(4),
      cd
    )
    original_rId <- gf1$chart_rId

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp), add = TRUE)
    prs$save(tmp)

    prs2   <- pptx_presentation(tmp)
    slide2 <- prs2$slides[[1]]
    shapes2 <- slide2$shapes$to_list()
    gf2    <- Filter(function(s) isTRUE(s$has_chart), shapes2)[[1]]

    expect_false(is.null(gf2$chart_rId))
    expect_true(grepl("^rId", gf2$chart_rId))
  })
})


# ============================================================================
# Helpers — build Chart domain objects from generated XML
# ============================================================================

# Helper: generate chartSpace XML for a category chart and return Chart object
.make_cat_chart <- function(chart_type = XL_CHART_TYPE$COLUMN_CLUSTERED,
                            categories = c("A", "B", "C"),
                            series     = list(list("S1", c(1, 2, 3)))) {
  cd <- CategoryChartData$new()
  cd$categories <- categories
  for (s in series) cd$add_series(s[[1]], s[[2]])
  xml_str <- chart_xml_writer(chart_type, cd)$xml
  node <- xml2::read_xml(xml_str)
  Chart$new(wrap_element(node), NULL)
}


# ============================================================================
# chart$chart_type — plot type inspection
# ============================================================================

describe("chart$chart_type", {
  it("returns COLUMN_CLUSTERED for a clustered column chart", {
    ch <- .make_cat_chart(XL_CHART_TYPE$COLUMN_CLUSTERED)
    expect_equal(ch$chart_type, XL_CHART_TYPE$COLUMN_CLUSTERED)
  })

  it("returns COLUMN_STACKED for a stacked column chart", {
    ch <- .make_cat_chart(XL_CHART_TYPE$COLUMN_STACKED)
    expect_equal(ch$chart_type, XL_CHART_TYPE$COLUMN_STACKED)
  })

  it("returns BAR_CLUSTERED for a horizontal bar chart", {
    ch <- .make_cat_chart(XL_CHART_TYPE$BAR_CLUSTERED)
    expect_equal(ch$chart_type, XL_CHART_TYPE$BAR_CLUSTERED)
  })

  it("returns LINE for a line chart without markers", {
    ch <- .make_cat_chart(XL_CHART_TYPE$LINE)
    expect_equal(ch$chart_type, XL_CHART_TYPE$LINE)
  })

  it("returns LINE_MARKERS for a line chart with markers", {
    ch <- .make_cat_chart(XL_CHART_TYPE$LINE_MARKERS)
    expect_equal(ch$chart_type, XL_CHART_TYPE$LINE_MARKERS)
  })

  it("returns PIE for a pie chart", {
    ch <- .make_cat_chart(XL_CHART_TYPE$PIE)
    expect_equal(ch$chart_type, XL_CHART_TYPE$PIE)
  })

  it("returns AREA for an area chart", {
    ch <- .make_cat_chart(XL_CHART_TYPE$AREA)
    expect_equal(ch$chart_type, XL_CHART_TYPE$AREA)
  })
})


# ============================================================================
# chart$series — SeriesCollection
# ============================================================================

describe("chart$series", {
  it("returns a SeriesCollection with correct length", {
    ch <- .make_cat_chart(
      series = list(list("Alpha", c(1, 2, 3)), list("Beta", c(4, 5, 6)))
    )
    sc <- ch$series
    expect_equal(length(sc), 2L)
  })

  it("series have correct index, name, and values", {
    ch <- .make_cat_chart(
      categories = c("Q1", "Q2"),
      series     = list(list("Revenue", c(100, 200)))
    )
    sc  <- ch$series
    ser <- sc[[1]]
    expect_equal(ser$index, 0L)
    expect_equal(ser$name, "Revenue")
    expect_equal(ser$values, c(100, 200))
  })

  it("series are ordered by c:order, not DOM order", {
    ch <- .make_cat_chart(
      series = list(
        list("First",  c(1, 2)),
        list("Second", c(3, 4)),
        list("Third",  c(5, 6))
      )
    )
    sc <- ch$series
    expect_equal(length(sc), 3L)
    expect_equal(sc[[1]]$name, "First")
    expect_equal(sc[[3]]$name, "Third")
  })
})


# ============================================================================
# chart$plots — plot collection
# ============================================================================

describe("chart$plots", {
  it("returns a list of one plot for a simple chart", {
    ch <- .make_cat_chart()
    plots <- ch$plots
    expect_equal(length(plots), 1L)
    expect_true(inherits(plots[[1]], "BarPlot"))
  })

  it("BarPlot has a vary_by_categories property", {
    ch <- .make_cat_chart(XL_CHART_TYPE$COLUMN_CLUSTERED)
    pl <- ch$plots[[1]]
    expect_true(inherits(pl, "BarPlot"))
    expect_false(is.null(pl$vary_by_categories))
  })

  it("BarPlot gap_width defaults to 150", {
    ch <- .make_cat_chart(XL_CHART_TYPE$COLUMN_CLUSTERED)
    pl <- ch$plots[[1]]
    expect_equal(pl$gap_width, 150L)
  })

  it("BarPlot gap_width setter changes the value", {
    ch <- .make_cat_chart(XL_CHART_TYPE$COLUMN_CLUSTERED)
    pl <- ch$plots[[1]]
    pl$gap_width <- 200L
    expect_equal(pl$gap_width, 200L)
  })

  it("BarPlot overlap setter changes the value", {
    ch <- .make_cat_chart(XL_CHART_TYPE$COLUMN_STACKED)
    pl <- ch$plots[[1]]
    pl$overlap <- 50L
    expect_equal(pl$overlap, 50L)
    pl$overlap <- 0L
    expect_equal(pl$overlap, 0L)
  })

  it("BarPlot vary_by_categories setter changes the value", {
    ch <- .make_cat_chart(XL_CHART_TYPE$COLUMN_CLUSTERED)
    pl <- ch$plots[[1]]
    pl$vary_by_categories <- TRUE
    expect_true(pl$vary_by_categories)
    pl$vary_by_categories <- FALSE
    expect_false(pl$vary_by_categories)
  })

  it("LineSeries smooth setter changes the value", {
    ch <- .make_cat_chart(XL_CHART_TYPE$LINE)
    ser <- ch$plots[[1]]$series[[1]]
    ser$smooth <- FALSE
    expect_false(ser$smooth)
    ser$smooth <- TRUE
    expect_true(ser$smooth)
  })

  it("BarSeries invert_if_negative setter changes the value", {
    ch <- .make_cat_chart(XL_CHART_TYPE$COLUMN_CLUSTERED)
    ser <- ch$plots[[1]]$series[[1]]
    ser$invert_if_negative <- FALSE
    expect_false(ser$invert_if_negative)
    ser$invert_if_negative <- TRUE
    expect_true(ser$invert_if_negative)
  })

  it("LinePlot is returned for a line chart", {
    ch <- .make_cat_chart(XL_CHART_TYPE$LINE)
    pl <- ch$plots[[1]]
    expect_true(inherits(pl, "LinePlot"))
  })

  it("PiePlot is returned for a pie chart", {
    ch <- .make_cat_chart(XL_CHART_TYPE$PIE)
    pl <- ch$plots[[1]]
    expect_true(inherits(pl, "PiePlot"))
  })
})


# ============================================================================
# chart$has_legend / chart$legend
# ============================================================================

describe("chart$has_legend", {
  it("returns FALSE for a chart with no legend", {
    ch <- .make_cat_chart()
    expect_false(ch$has_legend)
  })

  it("setting has_legend to TRUE adds a legend", {
    ch <- .make_cat_chart()
    ch$has_legend <- TRUE
    expect_true(ch$has_legend)
    expect_false(is.null(ch$legend))
  })

  it("setting has_legend to FALSE removes the legend", {
    ch <- .make_cat_chart()
    ch$has_legend <- TRUE
    ch$has_legend <- FALSE
    expect_false(ch$has_legend)
    expect_null(ch$legend)
  })
})


describe("Legend$position", {
  it("returns RIGHT (default) when no legendPos element exists", {
    ch <- .make_cat_chart()
    ch$has_legend <- TRUE
    lgnd <- ch$legend
    expect_equal(lgnd$position, XL_LEGEND_POSITION$RIGHT)
  })

  it("position can be set and read back", {
    ch <- .make_cat_chart()
    ch$has_legend <- TRUE
    lgnd <- ch$legend
    lgnd$position <- XL_LEGEND_POSITION$BOTTOM
    expect_equal(lgnd$position, XL_LEGEND_POSITION$BOTTOM)
  })
})


# ============================================================================
# chart$has_title
# ============================================================================

describe("chart$has_title", {
  it("returns FALSE when no title element", {
    ch <- .make_cat_chart()
    expect_false(ch$has_title)
  })

  it("setting has_title to TRUE creates a title element", {
    ch <- .make_cat_chart()
    ch$has_title <- TRUE
    expect_true(ch$has_title)
  })

  it("setting has_title to FALSE removes the title", {
    ch <- .make_cat_chart()
    ch$has_title <- TRUE
    ch$has_title <- FALSE
    expect_false(ch$has_title)
  })
})


# ============================================================================
# chart$category_axis
# ============================================================================

describe("chart$category_axis", {
  it("returns a CategoryAxis for category charts", {
    ch <- .make_cat_chart()
    ax <- ch$category_axis
    expect_true(inherits(ax, "CategoryAxis"))
  })

  it("category axis has_major_gridlines returns FALSE by default", {
    ch <- .make_cat_chart()
    ax <- ch$category_axis
    expect_false(ax$has_major_gridlines)
  })

  it("setting has_major_gridlines to TRUE works", {
    ch <- .make_cat_chart()
    ax <- ch$category_axis
    ax$has_major_gridlines <- TRUE
    expect_true(ax$has_major_gridlines)
  })

  it("tick_label_position defaults to nextTo for generated charts", {
    ch <- .make_cat_chart()
    ax <- ch$category_axis
    expect_equal(ax$tick_label_position, XL_TICK_LABEL_POSITION$NEXT_TO_AXIS)
  })

  it("tick_label_position can be set", {
    ch <- .make_cat_chart()
    ax <- ch$category_axis
    ax$tick_label_position <- XL_TICK_LABEL_POSITION$NONE
    expect_equal(ax$tick_label_position, XL_TICK_LABEL_POSITION$NONE)
  })

  it("category_type is CATEGORY_SCALE for CategoryAxis", {
    ch <- .make_cat_chart()
    ax <- ch$category_axis
    expect_equal(ax$category_type, XL_CATEGORY_TYPE$CATEGORY_SCALE)
  })

  it("reverse_order defaults to FALSE", {
    ch <- .make_cat_chart()
    ax <- ch$category_axis
    expect_false(ax$reverse_order)
  })

  it("reverse_order can be toggled", {
    ch <- .make_cat_chart()
    ax <- ch$category_axis
    ax$reverse_order <- TRUE
    expect_true(ax$reverse_order)
  })
})


# ============================================================================
# chart$value_axis
# ============================================================================

describe("chart$value_axis", {
  it("returns a ValueAxis for category charts", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    expect_true(inherits(ax, "ValueAxis"))
  })

  it("has_major_gridlines returns TRUE by default (standard xml writer output)", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    # xmlwriter adds c:majorGridlines by default
    expect_true(ax$has_major_gridlines)
  })

  it("maximum_scale defaults to NULL", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    expect_null(ax$maximum_scale)
  })

  it("minimum_scale defaults to NULL", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    expect_null(ax$minimum_scale)
  })

  it("maximum_scale can be set and read", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    ax$maximum_scale <- 100
    expect_equal(ax$maximum_scale, 100)
  })

  it("minimum_scale can be set and read", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    ax$minimum_scale <- 0
    expect_equal(ax$minimum_scale, 0)
  })

  it("visible defaults to TRUE", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    expect_true(ax$visible)
  })

  it("setting visible=FALSE marks axis as deleted", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    ax$visible <- FALSE
    expect_false(ax$visible)
  })

  it("tick_labels$number_format defaults to General", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    expect_equal(ax$tick_labels$number_format, "General")
  })

  it("tick_labels$number_format can be set and read back", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    ax$tick_labels$number_format <- "0.0%"
    expect_equal(ax$tick_labels$number_format, "0.0%")
  })

  it("tick_labels$number_format setter updates number_format_is_linked to FALSE", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    ax$tick_labels$number_format <- "#,##0"
    expect_false(ax$tick_labels$number_format_is_linked)
  })

  it("major_tick_mark can be set and read", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    ax$major_tick_mark <- XL_TICK_MARK$OUTSIDE
    expect_equal(ax$major_tick_mark, XL_TICK_MARK$OUTSIDE)
  })

  it("minor_tick_mark can be set and read", {
    ch <- .make_cat_chart()
    ax <- ch$value_axis
    ax$minor_tick_mark <- XL_TICK_MARK$INSIDE
    expect_equal(ax$minor_tick_mark, XL_TICK_MARK$INSIDE)
  })
})


# ============================================================================
# BarPlot$has_data_labels / DataLabels
# ============================================================================

describe("BasePlot$has_data_labels", {
  it("returns FALSE when no dLbls element", {
    ch <- .make_cat_chart()
    pl <- ch$plots[[1]]
    expect_false(pl$has_data_labels)
  })

  it("setting has_data_labels to TRUE adds dLbls", {
    ch <- .make_cat_chart()
    pl <- ch$plots[[1]]
    pl$has_data_labels <- TRUE
    expect_true(pl$has_data_labels)
  })

  it("data_labels returns a DataLabels object after enabling", {
    ch <- .make_cat_chart()
    pl <- ch$plots[[1]]
    pl$has_data_labels <- TRUE
    dl <- pl$data_labels
    expect_true(inherits(dl, "DataLabels"))
  })

  it("setting has_data_labels to FALSE removes dLbls", {
    ch <- .make_cat_chart()
    pl <- ch$plots[[1]]
    pl$has_data_labels <- TRUE
    pl$has_data_labels <- FALSE
    expect_false(pl$has_data_labels)
  })
})


# ============================================================================
# New enum constants
# ============================================================================

describe("XL_AXIS_CROSSES constants", {
  it("contains expected XML string values", {
    expect_equal(XL_AXIS_CROSSES$AUTOMATIC, "autoZero")
    expect_equal(XL_AXIS_CROSSES$MAXIMUM,   "max")
    expect_equal(XL_AXIS_CROSSES$MINIMUM,   "min")
  })
})

describe("XL_LEGEND_POSITION constants", {
  it("contains expected XML string values", {
    expect_equal(XL_LEGEND_POSITION$BOTTOM, "b")
    expect_equal(XL_LEGEND_POSITION$RIGHT,  "r")
    expect_equal(XL_LEGEND_POSITION$TOP,    "t")
  })
})

describe("XL_TICK_MARK constants", {
  it("contains expected XML string values", {
    expect_equal(XL_TICK_MARK$CROSS,   "cross")
    expect_equal(XL_TICK_MARK$INSIDE,  "in")
    expect_equal(XL_TICK_MARK$NONE,    "none")
    expect_equal(XL_TICK_MARK$OUTSIDE, "out")
  })
})

describe("XL_TICK_LABEL_POSITION constants", {
  it("contains expected values", {
    expect_equal(XL_TICK_LABEL_POSITION$HIGH,         "high")
    expect_equal(XL_TICK_LABEL_POSITION$NEXT_TO_AXIS, "nextTo")
    expect_equal(XL_TICK_LABEL_POSITION$NONE,         "none")
  })
})

describe("XL_DATA_LABEL_POSITION / XL_LABEL_POSITION", {
  it("has expected values", {
    expect_equal(XL_DATA_LABEL_POSITION$CENTER,      "ctr")
    expect_equal(XL_DATA_LABEL_POSITION$OUTSIDE_END, "outEnd")
  })

  it("XL_LABEL_POSITION is an alias for XL_DATA_LABEL_POSITION", {
    expect_identical(XL_LABEL_POSITION, XL_DATA_LABEL_POSITION)
  })
})


describe("Bubble / Radar / Doughnut chart round-trips", {
  make_slide <- function() {
    prs <- pptx_presentation()
    list(prs = prs, slide = prs$slides$add_slide(prs$slide_layouts[[6]]))
  }

  it("bubble chart saves and reloads", {
    e <- make_slide()
    bd <- BubbleChartData$new()
    s  <- bd$add_series("B1")
    s$add_data_point(1, 2, 10)
    s$add_data_point(2, 3, 20)
    e$slide$shapes$add_chart(
      XL_CHART_TYPE$BUBBLE, Inches(1), Inches(1), Inches(4), Inches(3), bd
    )
    tmp <- tempfile(fileext = ".pptx")
    e$prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    gfs  <- Filter(function(s) inherits(s, "GraphicFrame"),
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(length(gfs), 1L)
    expect_true(gfs[[1]]$has_chart)
  })

  it("radar chart saves and reloads", {
    e <- make_slide()
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B", "C")
    cd$add_series("R1", c(1, 2, 3))
    e$slide$shapes$add_chart(
      XL_CHART_TYPE$RADAR, Inches(1), Inches(1), Inches(4), Inches(3), cd
    )
    tmp <- tempfile(fileext = ".pptx")
    e$prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    gfs  <- Filter(function(s) inherits(s, "GraphicFrame"),
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(length(gfs), 1L)
  })

  it("doughnut chart saves and reloads", {
    e <- make_slide()
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B", "C")
    cd$add_series("D1", c(40, 35, 25))
    e$slide$shapes$add_chart(
      XL_CHART_TYPE$DOUGHNUT, Inches(1), Inches(1), Inches(4), Inches(3), cd
    )
    tmp <- tempfile(fileext = ".pptx")
    e$prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    gfs  <- Filter(function(s) inherits(s, "GraphicFrame"),
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(length(gfs), 1L)
  })

  it("line chart saves and reloads with correct type", {
    e <- make_slide()
    cd <- CategoryChartData$new()
    cd$categories <- c("Jan", "Feb", "Mar")
    cd$add_series("Sales", c(100, 120, 115))
    e$slide$shapes$add_chart(
      XL_CHART_TYPE$LINE, Inches(1), Inches(1), Inches(4), Inches(3), cd
    )
    tmp <- tempfile(fileext = ".pptx")
    e$prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    gfs  <- Filter(function(s) inherits(s, "GraphicFrame"),
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(length(gfs), 1L)
    expect_equal(gfs[[1]]$chart$chart_type, XL_CHART_TYPE$LINE)
  })

  it("pie chart saves and reloads with correct type", {
    e <- make_slide()
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B", "C")
    cd$add_series("Share", c(50, 30, 20))
    e$slide$shapes$add_chart(
      XL_CHART_TYPE$PIE, Inches(1), Inches(1), Inches(4), Inches(3), cd
    )
    tmp <- tempfile(fileext = ".pptx")
    e$prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    gfs  <- Filter(function(s) inherits(s, "GraphicFrame"),
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(length(gfs), 1L)
    expect_equal(gfs[[1]]$chart$chart_type, XL_CHART_TYPE$PIE)
  })

  it("area chart saves and reloads with correct type", {
    e <- make_slide()
    cd <- CategoryChartData$new()
    cd$categories <- c("Q1", "Q2", "Q3")
    cd$add_series("Revenue", c(400, 500, 450))
    e$slide$shapes$add_chart(
      XL_CHART_TYPE$AREA, Inches(1), Inches(1), Inches(4), Inches(3), cd
    )
    tmp <- tempfile(fileext = ".pptx")
    e$prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    gfs  <- Filter(function(s) inherits(s, "GraphicFrame"),
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(length(gfs), 1L)
    expect_equal(gfs[[1]]$chart$chart_type, XL_CHART_TYPE$AREA)
  })

  it("XY scatter chart saves and reloads", {
    e  <- make_slide()
    xd <- XyChartData$new()
    s  <- xd$add_series("Series 1")
    s$add_data_point(1.0, 2.0)
    s$add_data_point(2.0, 3.5)
    s$add_data_point(3.0, 2.8)
    e$slide$shapes$add_chart(
      XL_CHART_TYPE$XY_SCATTER, Inches(1), Inches(1), Inches(4), Inches(3), xd
    )
    tmp <- tempfile(fileext = ".pptx")
    e$prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    gfs  <- Filter(function(s) inherits(s, "GraphicFrame"),
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(length(gfs), 1L)
    expect_true(gfs[[1]]$has_chart)
  })

  it("bar clustered chart saves and reloads with correct type", {
    e <- make_slide()
    cd <- CategoryChartData$new()
    cd$categories <- c("North", "South", "East")
    cd$add_series("Units", c(30, 45, 25))
    e$slide$shapes$add_chart(
      XL_CHART_TYPE$BAR_CLUSTERED, Inches(1), Inches(1), Inches(4), Inches(3), cd
    )
    tmp <- tempfile(fileext = ".pptx")
    e$prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    gfs  <- Filter(function(s) inherits(s, "GraphicFrame"),
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(length(gfs), 1L)
    expect_equal(gfs[[1]]$chart$chart_type, XL_CHART_TYPE$BAR_CLUSTERED)
  })
})
