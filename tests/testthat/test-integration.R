# Integration / round-trip tests.
#
# These tests build a multi-feature presentation in memory, save it to a
# temp file, reload it, and verify that every feature survived the round-trip.
# This catches serialization bugs that unit tests on individual objects miss.

library(rpptx)

pptx_path <- function(name) system.file("test_files", name, package = "rpptx")

# ============================================================================
# Full round-trip: shapes, text, fill, line
# ============================================================================

describe("Integration: shapes + text + fill + line round-trip", {
  it("all properties survive a save/load cycle", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[6]]  # blank layout
    slide  <- prs$slides$add_slide(layout)

    # Rectangle with solid fill and line
    rect <- slide$shapes$add_shape(
      MSO_AUTO_SHAPE_TYPE$RECTANGLE,
      Inches(0.5), Inches(0.5), Inches(3), Inches(2)
    )
    rect$name <- "MyRect"
    rect$fill$solid()
    rect$fill$fore_color$rgb <- RGBColor(0x4F, 0x81, 0xBD)
    rect$line$width <- Pt(2)
    rect$line$color$rgb <- RGBColor(0xFF, 0, 0)

    # Textbox
    txb <- slide$shapes$add_textbox(
      Inches(4), Inches(0.5), Inches(5), Inches(2)
    )
    tf <- txb$text_frame
    tf$word_wrap <- TRUE
    para <- tf$paragraphs[[1]]
    run  <- para$add_run()
    run$text      <- "Hello rpptx"
    run$font$bold <- TRUE
    run$font$size <- Pt(18)
    run$font$name <- "Arial"

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2  <- pptx_presentation(tmp)
    slide2 <- prs2$slides[[1]]

    # Find rect and textbox by iterating shapes
    shapes_list <- slide2$shapes$to_list()
    rect2 <- Filter(function(s) s$name == "MyRect", shapes_list)[[1]]
    txbs2 <- Filter(function(s) inherits(s, "Shape") && s$has_text_frame &&
                                  !isTRUE(s$is_placeholder),
                    shapes_list)
    txb2  <- txbs2[[length(txbs2)]]  # last added textbox

    # Shape fill
    expect_equal(rect2$fill$type, MSO_FILL$SOLID)
    expect_equal(rect2$fill$fore_color$rgb, RGBColor(0x4F, 0x81, 0xBD))

    # Shape line
    expect_equal(as.integer(rect2$line$width), as.integer(Pt(2)))
    expect_equal(rect2$line$color$rgb, RGBColor(0xFF, 0, 0))

    # Text
    r2 <- txb2$text_frame$paragraphs[[1]]$runs[[1]]
    expect_equal(r2$text, "Hello rpptx")
    expect_true(r2$font$bold)
    expect_equal(as.integer(r2$font$size), as.integer(Pt(18)))
    expect_equal(r2$font$name, "Arial")
  })
})


# ============================================================================
# Integration: table round-trip
# ============================================================================

describe("Integration: table round-trip", {
  it("table cell text and fill survive save/load", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[6]]
    slide  <- prs$slides$add_slide(layout)
    tbl_shape <- slide$shapes$add_table(
      3L, 3L, Inches(1), Inches(1), Inches(7), Inches(3))
    tbl <- tbl_shape$table

    c11 <- tbl$cell(1, 1); c11$text <- "Header A"
    c12 <- tbl$cell(1, 2); c12$text <- "Header B"
    c21 <- tbl$cell(2, 1); c21$text <- "Data 1"
    c22 <- tbl$cell(2, 2); c22$text <- "Data 2"
    c11$fill$solid()
    c11$fill$fore_color$rgb <- RGBColor(0x4F, 0x81, 0xBD)

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2  <- pptx_presentation(tmp)
    slide2 <- prs2$slides[[1]]
    tbl2_shape <- Filter(function(s) inherits(s, "GraphicFrame"),
                         slide2$shapes$to_list())[[1]]
    tbl2 <- tbl2_shape$table

    expect_equal(tbl2$cell(1, 1)$text, "Header A")
    expect_equal(tbl2$cell(1, 2)$text, "Header B")
    expect_equal(tbl2$cell(2, 1)$text, "Data 1")
    expect_equal(tbl2$cell(1, 1)$fill$type, MSO_FILL$SOLID)
    expect_equal(tbl2$cell(1, 1)$fill$fore_color$rgb,
                 RGBColor(0x4F, 0x81, 0xBD))
  })
})


# ============================================================================
# Integration: notes slide round-trip
# ============================================================================

describe("Integration: notes slide round-trip", {
  it("notes text survives save/load", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[6]]
    slide  <- prs$slides$add_slide(layout)
    slide$notes_slide$notes_text_frame$text <- "Speaker notes here."

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2   <- pptx_presentation(tmp)
    slide2 <- prs2$slides[[1]]

    expect_true(slide2$has_notes_slide)
    expect_equal(slide2$notes_slide$notes_text_frame$text, "Speaker notes here.")
  })
})


# ============================================================================
# Integration: slide operations round-trip
# ============================================================================

describe("Integration: slide reorder / delete round-trip", {
  it("slide order persists after move + save/load", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[6]]
    s1 <- prs$slides$add_slide(layout)
    s2 <- prs$slides$add_slide(layout)
    s3 <- prs$slides$add_slide(layout)
    txb1 <- s1$shapes$add_textbox(Inches(0), Inches(0), Inches(1), Inches(1))
    txb1$text_frame$text <- "SLIDE_1"
    txb2 <- s2$shapes$add_textbox(Inches(0), Inches(0), Inches(1), Inches(1))
    txb2$text_frame$text <- "SLIDE_2"
    txb3 <- s3$shapes$add_textbox(Inches(0), Inches(0), Inches(1), Inches(1))
    txb3$text_frame$text <- "SLIDE_3"

    # Move slide 3 to position 1
    prs$slides$move(s3, 1L)

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)

    # Helper to get last textbox text on a slide
    slide_label <- function(sl) {
      txbs <- Filter(function(s) inherits(s, "Shape") && s$has_text_frame,
                     sl$shapes$to_list())
      if (length(txbs) == 0) return(NA_character_)
      txbs[[length(txbs)]]$text_frame$text
    }

    expect_equal(slide_label(prs2$slides[[1]]), "SLIDE_3")
    expect_equal(slide_label(prs2$slides[[2]]), "SLIDE_1")
    expect_equal(slide_label(prs2$slides[[3]]), "SLIDE_2")
  })

  it("deleted slide is absent after save/load", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[6]]
    s1 <- prs$slides$add_slide(layout)
    s2 <- prs$slides$add_slide(layout)

    txb1 <- s1$shapes$add_textbox(Inches(0), Inches(0), Inches(1), Inches(1))
    txb1$text_frame$text <- "KEEP"
    txb2 <- s2$shapes$add_textbox(Inches(0), Inches(0), Inches(1), Inches(1))
    txb2$text_frame$text <- "DELETE"

    prs$slides$delete(s2)

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)

    expect_equal(length(prs2$slides), 1L)
    txbs <- Filter(function(s) inherits(s, "Shape") && s$has_text_frame,
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(txbs[[length(txbs)]]$text_frame$text, "KEEP")
  })
})


# ============================================================================
# Integration: multi-slide presentation with charts
# ============================================================================

describe("Integration: multi-slide presentation with chart", {
  it("chart and text on separate slides both survive save/load", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[6]]

    # Slide 1: text
    s1 <- prs$slides$add_slide(layout)
    txb_s1 <- s1$shapes$add_textbox(Inches(1), Inches(1), Inches(8), Inches(2))
    txb_s1$text_frame$text <- "Title slide"

    # Slide 2: chart
    s2 <- prs$slides$add_slide(layout)
    cd <- CategoryChartData$new()
    cd$categories <- c("A", "B", "C")
    cd$add_series("S1", c(10, 20, 30))
    s2$shapes$add_chart(
      XL_CHART_TYPE$COLUMN_CLUSTERED,
      Inches(1), Inches(1), Inches(8), Inches(5),
      cd
    )

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)

    expect_equal(length(prs2$slides), 2L)

    # Text on slide 1
    txbs <- Filter(function(s) inherits(s, "Shape") && s$has_text_frame,
                   prs2$slides[[1]]$shapes$to_list())
    expect_equal(txbs[[length(txbs)]]$text_frame$text, "Title slide")

    # Chart on slide 2
    gfs <- Filter(function(s) inherits(s, "GraphicFrame"),
                  prs2$slides[[2]]$shapes$to_list())
    expect_equal(length(gfs), 1L)
    expect_equal(gfs[[1]]$chart$chart_type, XL_CHART_TYPE$COLUMN_CLUSTERED)
  })
})


# ============================================================================
# Integration: core properties round-trip
# ============================================================================

describe("Integration: core properties round-trip", {
  it("title, author, and subject survive save/load", {
    prs <- pptx_presentation()
    prs$core_properties$title   <- "Integration Test"
    prs$core_properties$author  <- "rpptx"
    prs$core_properties$subject <- "Testing"

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)

    expect_equal(prs2$core_properties$title,   "Integration Test")
    expect_equal(prs2$core_properties$author,  "rpptx")
    expect_equal(prs2$core_properties$subject, "Testing")
  })
})


# ============================================================================
# Integration: open a real .pptx, mutate, save, verify
# ============================================================================

describe("Integration: mutate existing .pptx file", {
  it("can open test.pptx, add a slide, save, and reload", {
    src <- pptx_path("test.pptx")
    if (!file.exists(src)) skip("test.pptx fixture not found")

    prs    <- pptx_presentation(src)
    n_orig <- length(prs$slides)
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    txb <- slide$shapes$add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    txb$text_frame$text <- "ADDED"

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)

    expect_equal(length(prs2$slides), n_orig + 1L)
    new_slide <- prs2$slides[[n_orig + 1L]]
    txbs <- Filter(function(s) inherits(s, "Shape") && s$has_text_frame,
                   new_slide$shapes$to_list())
    expect_equal(txbs[[length(txbs)]]$text_frame$text, "ADDED")
  })
})
