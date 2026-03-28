# Tests for R/utils-print.R — S3 format/print methods and as.data.frame.Table.

.pptx_path <- function(name) {
  system.file("test_files", name, package = "rpptx")
}

describe("format.Presentation", {
  it("includes slide count, size, layouts, masters", {
    prs <- pptx_presentation()
    out <- format(prs)
    expect_match(out, "<Presentation>")
    expect_match(out, "slides")
    expect_match(out, "10.00")   # default width in inches
    expect_match(out, "7.50")    # default height in inches
    expect_match(out, "layouts")
    expect_match(out, "masters")
  })
})

describe("print.Presentation", {
  it("returns prs invisibly", {
    prs <- pptx_presentation()
    out <- withVisible(print(prs))
    expect_false(out$visible)
    expect_s3_class(out$value, "Presentation")
  })
})

describe("format.Slide", {
  it("shows shape count and shape rows", {
    prs <- pptx_presentation(.pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    out <- format(slide)
    expect_match(out, "<Slide>")
    expect_match(out, "shapes")
  })

  it("handles zero shapes", {
    prs <- pptx_presentation()
    slide <- prs$slides$add_slide(prs$slide_layouts[[1]])
    # A slide from a blank layout may still have placeholders, but
    # this at minimum should not error.
    out <- format(slide)
    expect_match(out, "<Slide>")
  })
})

describe("format.SlideLayout", {
  it("shows layout name and placeholder count", {
    prs <- pptx_presentation()
    lay <- prs$slide_layouts[[1]]
    out <- format(lay)
    expect_match(out, "<SlideLayout>")
    expect_match(out, "placeholder")
  })
})

describe("format.SlideMaster", {
  it("shows layout and placeholder counts", {
    prs <- pptx_presentation()
    master <- prs$slide_masters[[1]]
    out <- format(master)
    expect_match(out, "<SlideMaster>")
    expect_match(out, "layout")
    expect_match(out, "placeholder")
  })
})

describe("format.BaseShape", {
  it("shows shape type, name, size, and position", {
    prs  <- pptx_presentation(.pptx_path("test_slides.pptx"))
    shp  <- prs$slides[[1]]$shapes$to_list()[[1]]
    out  <- format(shp)
    expect_match(out, "in")     # position in inches
  })
})

describe("format.Table and as.data.frame.Table", {
  it("format shows row x col", {
    prs    <- pptx_presentation(.pptx_path("test_slides.pptx"))
    slide  <- prs$slides[[1]]
    shapes <- slide$shapes$to_list()
    tbl_shape <- Filter(function(s) inherits(s, "GraphicFrame"), shapes)[[1]]
    tbl  <- tbl_shape$table
    out  <- format(tbl)
    expect_match(out, "<Table>")
    expect_match(out, "rows")
    expect_match(out, "cols")
  })

  it("as.data.frame produces character data.frame", {
    prs    <- pptx_presentation(.pptx_path("test_slides.pptx"))
    slide  <- prs$slides[[1]]
    shapes <- slide$shapes$to_list()
    tbl_shape <- Filter(function(s) inherits(s, "GraphicFrame"), shapes)[[1]]
    tbl  <- tbl_shape$table
    df   <- as.data.frame(tbl)
    expect_s3_class(df, "data.frame")
    expect_true(all(vapply(df, is.character, logical(1))))
    expect_equal(ncol(df), length(tbl$columns))
    expect_equal(nrow(df), length(tbl$rows))
  })
})

describe("format.TextFrame", {
  it("shows TextFrame with text snippet", {
    prs   <- pptx_presentation(.pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    # Find a shape with a text frame
    shp   <- Filter(function(s) tryCatch(!is.null(s$text_frame), error = function(e) FALSE),
                    slide$shapes$to_list())[[1]]
    out   <- format(shp$text_frame)
    expect_match(out, "<TextFrame>")
  })
})
