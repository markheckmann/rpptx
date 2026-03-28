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
