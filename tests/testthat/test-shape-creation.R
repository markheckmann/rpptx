# Tests for Phase 6: shape creation — add_textbox, add_shape,
# clone_layout_placeholders, iter_cloneable_placeholders.

pptx_path <- function(name) system.file("test_files", name, package = "rpptx")

# Helper: fresh presentation with one blank-ish slide (no placeholders)
new_blank_slide <- function() {
  prs    <- pptx_presentation(pptx_path("no-slides.pptx"))
  layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
  prs$slides$add_slide(layout)
}


# ============================================================================
# MSO_AUTO_SHAPE_TYPE
# ============================================================================

describe("MSO_AUTO_SHAPE_TYPE", {
  it("has RECTANGLE with correct integer and prst", {
    r <- MSO_AUTO_SHAPE_TYPE$RECTANGLE
    expect_equal(r$value, 1L)
    expect_equal(r$prst,  "rect")
  })

  it("has ROUNDED_RECTANGLE with correct prst", {
    expect_equal(MSO_AUTO_SHAPE_TYPE$ROUNDED_RECTANGLE$prst, "roundRect")
  })

  it("has OVAL with correct prst", {
    expect_equal(MSO_AUTO_SHAPE_TYPE$OVAL$prst, "ellipse")
  })

  it("MSO_SHAPE alias equals MSO_AUTO_SHAPE_TYPE", {
    expect_identical(MSO_SHAPE, MSO_AUTO_SHAPE_TYPE)
  })

  it("has expected members covering common shapes", {
    expected <- c("RECTANGLE", "ROUNDED_RECTANGLE", "OVAL", "DIAMOND",
                  "ISOSCELES_TRIANGLE", "RIGHT_TRIANGLE", "HEXAGON",
                  "STAR_5_POINT", "HEART", "CROSS")
    for (nm in expected) {
      expect_true(!is.null(MSO_AUTO_SHAPE_TYPE[[nm]]),
                  info = sprintf("MSO_AUTO_SHAPE_TYPE$%s should exist", nm))
    }
  })
})


# ============================================================================
# CT_Shape factory functions
# ============================================================================

describe("CT_Shape_new_textbox_sp()", {
  sp <- CT_Shape_new_textbox_sp(3L, "TextBox 2", 100L, 200L, 300L, 400L)

  it("returns a CT_Shape", {
    expect_true(inherits(sp, "CT_Shape"))
  })

  it("has correct shape id", {
    expect_equal(sp$shape_id, 3L)
  })

  it("has correct shape name", {
    expect_equal(sp$shape_name, "TextBox 2")
  })

  it("has correct position", {
    expect_equal(as.integer(sp$x), 100L)
    expect_equal(as.integer(sp$y), 200L)
  })

  it("has correct dimensions", {
    expect_equal(as.integer(sp$cx), 300L)
    expect_equal(as.integer(sp$cy), 400L)
  })

  it("has txBox=1 attribute on cNvSpPr", {
    nodes <- sp$xpath("./*[1]/p:cNvSpPr")
    expect_equal(length(nodes), 1L)
    expect_equal(xml2::xml_attr(nodes[[1]], "txBox"), "1")
  })

  it("has p:txBody with bodyPr wrap=none", {
    nodes <- sp$xpath("p:txBody/a:bodyPr")
    expect_equal(length(nodes), 1L)
    expect_equal(xml2::xml_attr(nodes[[1]], "wrap"), "none")
  })
})


describe("CT_Shape_new_autoshape_sp()", {
  sp <- CT_Shape_new_autoshape_sp(5L, "AutoShape 4", "ellipse",
                                  Inches(1), Inches(2), Inches(3), Inches(4))

  it("returns a CT_Shape", {
    expect_true(inherits(sp, "CT_Shape"))
  })

  it("has correct shape id and name", {
    expect_equal(sp$shape_id, 5L)
    expect_equal(sp$shape_name, "AutoShape 4")
  })

  it("has prstGeom with correct prst attribute", {
    nodes <- sp$xpath("p:spPr/a:prstGeom")
    expect_equal(length(nodes), 1L)
    expect_equal(xml2::xml_attr(nodes[[1]], "prst"), "ellipse")
  })

  it("has p:style child", {
    nodes <- sp$xpath("p:style")
    expect_equal(length(nodes), 1L)
  })

  it("has p:txBody child", {
    nodes <- sp$xpath("p:txBody")
    expect_equal(length(nodes), 1L)
  })

  it("is not a placeholder", {
    expect_false(sp$has_ph_elm)
  })
})


describe("CT_Shape_new_placeholder_sp()", {
  sp <- CT_Shape_new_placeholder_sp(
    id      = 2L,
    name    = "Title 1",
    ph_type = "title",
    orient  = "horz",
    sz      = "full",
    idx     = 0L
  )

  it("returns a CT_Shape", {
    expect_true(inherits(sp, "CT_Shape"))
  })

  it("is a placeholder", {
    expect_true(sp$has_ph_elm)
  })

  it("has correct ph type", {
    expect_equal(sp$ph_type, "title")
  })

  it("has correct ph idx", {
    expect_equal(sp$ph_idx, 0L)
  })

  it("has txBody for text-bearing type", {
    nodes <- sp$xpath("p:txBody")
    expect_equal(length(nodes), 1L)
  })

  it("does NOT add txBody for non-text type", {
    sp_pic <- CT_Shape_new_placeholder_sp(3L, "Picture 2", "pic", "horz", "full", 2L)
    nodes  <- sp_pic$xpath("p:txBody")
    expect_equal(length(nodes), 0L)
  })
})


# ============================================================================
# CT_GroupShape$add_textbox / add_autoshape / add_placeholder
# ============================================================================

describe("CT_GroupShape$add_textbox()", {
  it("adds a textbox shape to the spTree", {
    prs   <- pptx_presentation(pptx_path("no-slides.pptx"))
    slide <- prs$slides$add_slide(prs$slide_masters[[1]]$slide_layouts[[1]])
    spTree <- slide$element$spTree
    n_before <- length(spTree$iter_shape_elms())
    sp <- spTree$add_textbox(99L, "TextBox 98", 100L, 200L, 300L, 400L)
    expect_equal(length(spTree$iter_shape_elms()), n_before + 1L)
    expect_true(inherits(sp, "CT_Shape"))
    expect_equal(sp$shape_id, 99L)
  })
})

describe("CT_GroupShape$add_autoshape()", {
  it("adds an autoshape to the spTree", {
    prs   <- pptx_presentation(pptx_path("no-slides.pptx"))
    slide <- prs$slides$add_slide(prs$slide_masters[[1]]$slide_layouts[[1]])
    spTree <- slide$element$spTree
    n_before <- length(spTree$iter_shape_elms())
    sp <- spTree$add_autoshape(100L, "AutoShape 99", "rect",
                               Inches(1), Inches(1), Inches(2), Inches(2))
    expect_equal(length(spTree$iter_shape_elms()), n_before + 1L)
    expect_true(inherits(sp, "CT_Shape"))
  })
})


# ============================================================================
# SlideShapes$add_textbox()
# ============================================================================

describe("SlideShapes$add_textbox()", {
  it("returns a Shape wrapping a textbox sp", {
    slide  <- new_blank_slide()
    shape  <- slide$shapes$add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
    expect_true(inherits(shape, "Shape"))
  })

  it("shape has has_text_frame TRUE", {
    slide  <- new_blank_slide()
    shape  <- slide$shapes$add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
    expect_true(shape$has_text_frame)
  })

  it("shape has correct position and size", {
    slide  <- new_blank_slide()
    shape  <- slide$shapes$add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
    expect_equal(as.integer(shape$left),   as.integer(Inches(1)))
    expect_equal(as.integer(shape$top),    as.integer(Inches(2)))
    expect_equal(as.integer(shape$width),  as.integer(Inches(3)))
    expect_equal(as.integer(shape$height), as.integer(Inches(1)))
  })

  it("text_frame is accessible and writable", {
    slide  <- new_blank_slide()
    shape  <- slide$shapes$add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    shape$text_frame$text <- "hello"
    expect_equal(shape$text_frame$text, "hello")
  })

  it("increments slide shape count", {
    slide  <- new_blank_slide()
    before <- length(slide$shapes)
    slide$shapes$add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
    expect_equal(length(slide$shapes), before + 1L)
  })

  it("name follows 'TextBox N' pattern", {
    slide  <- new_blank_slide()
    shape  <- slide$shapes$add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
    expect_match(shape$name, "^TextBox \\d+$")
  })
})


# ============================================================================
# SlideShapes$add_shape()
# ============================================================================

describe("SlideShapes$add_shape()", {
  it("returns a Shape", {
    slide <- new_blank_slide()
    shape <- slide$shapes$add_shape(
      MSO_AUTO_SHAPE_TYPE$RECTANGLE,
      Inches(1), Inches(1), Inches(2), Inches(1)
    )
    expect_true(inherits(shape, "Shape"))
  })

  it("shape has correct prst geometry", {
    slide <- new_blank_slide()
    shape <- slide$shapes$add_shape(
      MSO_AUTO_SHAPE_TYPE$OVAL,
      Inches(1), Inches(1), Inches(2), Inches(2)
    )
    nodes <- shape$element$xpath("p:spPr/a:prstGeom")
    expect_equal(xml2::xml_attr(nodes[[1]], "prst"), "ellipse")
  })

  it("shape has correct position", {
    slide <- new_blank_slide()
    shape <- slide$shapes$add_shape(
      MSO_AUTO_SHAPE_TYPE$RECTANGLE,
      Inches(2), Inches(3), Inches(4), Inches(1)
    )
    expect_equal(as.integer(shape$left),   as.integer(Inches(2)))
    expect_equal(as.integer(shape$top),    as.integer(Inches(3)))
    expect_equal(as.integer(shape$width),  as.integer(Inches(4)))
    expect_equal(as.integer(shape$height), as.integer(Inches(1)))
  })

  it("increments shape count", {
    slide  <- new_blank_slide()
    before <- length(slide$shapes)
    slide$shapes$add_shape(MSO_AUTO_SHAPE_TYPE$DIAMOND,
                           Inches(1), Inches(1), Inches(2), Inches(2))
    expect_equal(length(slide$shapes), before + 1L)
  })

  it("shape has has_text_frame TRUE (autoshapes always have txBody)", {
    slide <- new_blank_slide()
    shape <- slide$shapes$add_shape(
      MSO_AUTO_SHAPE_TYPE$RECTANGLE,
      Inches(1), Inches(1), Inches(2), Inches(2)
    )
    expect_true(shape$has_text_frame)
  })

  it("errors when passed a non-MSO_AUTO_SHAPE_TYPE value", {
    slide <- new_blank_slide()
    expect_error(
      slide$shapes$add_shape("rect", Inches(1), Inches(1), Inches(2), Inches(2)),
      "MSO_AUTO_SHAPE_TYPE"
    )
  })
})


# ============================================================================
# SlideLayout$iter_cloneable_placeholders()
# ============================================================================

describe("SlideLayout$iter_cloneable_placeholders()", {
  it("returns a list", {
    prs    <- pptx_presentation(pptx_path("test_slides.pptx"))
    layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
    result <- layout$iter_cloneable_placeholders()
    expect_true(is.list(result))
  })

  it("excludes date placeholder type", {
    prs    <- pptx_presentation(pptx_path("test_slides.pptx"))
    layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
    phs    <- layout$iter_cloneable_placeholders()
    types  <- vapply(phs, function(e) e$ph_type, character(1))
    expect_false("dt" %in% types)
  })

  it("excludes footer placeholder type", {
    prs    <- pptx_presentation(pptx_path("test_slides.pptx"))
    layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
    phs    <- layout$iter_cloneable_placeholders()
    types  <- vapply(phs, function(e) e$ph_type, character(1))
    expect_false("ftr" %in% types)
  })

  it("excludes slide number placeholder type", {
    prs    <- pptx_presentation(pptx_path("test_slides.pptx"))
    layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
    phs    <- layout$iter_cloneable_placeholders()
    types  <- vapply(phs, function(e) e$ph_type, character(1))
    expect_false("sldNum" %in% types)
  })
})


# ============================================================================
# SlideShapes$clone_layout_placeholders()
# ============================================================================

describe("SlideShapes$clone_layout_placeholders()", {
  it("adds placeholders from layout to slide (not just no-op)", {
    prs    <- pptx_presentation(pptx_path("test_slides.pptx"))
    layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
    cloneable <- layout$iter_cloneable_placeholders()
    # Only meaningful test if layout has cloneable placeholders
    skip_if(length(cloneable) == 0L, "layout has no cloneable placeholders")

    # Add a new slide and count its placeholders
    slide  <- prs$slides$add_slide(layout)
    phs    <- slide$shapes$to_list()
    ph_shapes <- Filter(function(s) s$is_placeholder, phs)
    expect_true(length(ph_shapes) >= length(cloneable))
  })

  it("cloned placeholders are Shape instances", {
    prs    <- pptx_presentation(pptx_path("test_slides.pptx"))
    layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
    cloneable <- layout$iter_cloneable_placeholders()
    skip_if(length(cloneable) == 0L, "layout has no cloneable placeholders")

    slide  <- prs$slides$add_slide(layout)
    ph_shapes <- Filter(function(s) s$is_placeholder, slide$shapes$to_list())
    for (s in ph_shapes) {
      expect_true(inherits(s, "Shape") || inherits(s, "BaseShape"))
    }
  })
})


# ============================================================================
# add_picture
# ============================================================================

describe("SlideShapes$add_picture", {
  make_png <- function(width = 100L, height = 80L) {
    skip_if_not_installed("magick")
    tmp <- tempfile(fileext = ".png")
    magick::image_write(magick::image_blank(width, height, "red"), tmp)
    tmp
  }

  it("returns a Picture proxy", {
    f   <- make_png()
    prs <- pptx_presentation()
    sl  <- invisible(prs$slides$add_slide(prs$slide_layouts[[6]]))
    pic <- sl$shapes$add_picture(f, Inches(1), Inches(1))
    expect_true(inherits(pic, "Picture"))
  })

  it("uses native dimensions when width and height omitted", {
    f   <- make_png(200L, 100L)
    prs <- pptx_presentation()
    sl  <- invisible(prs$slides$add_slide(prs$slide_layouts[[6]]))
    pic <- sl$shapes$add_picture(f, Inches(0), Inches(0))
    # 200px @ 72dpi = 200/72 * 914400 EMU
    expect_equal(pic$width,  as.integer(round(914400 * 200 / 72)))
    expect_equal(pic$height, as.integer(round(914400 * 100 / 72)))
  })

  it("preserves aspect ratio when only width given", {
    f   <- make_png(200L, 100L)  # 2:1 ratio
    prs <- pptx_presentation()
    sl  <- invisible(prs$slides$add_slide(prs$slide_layouts[[6]]))
    pic <- sl$shapes$add_picture(f, Inches(0), Inches(0), width = Inches(4))
    expect_equal(pic$width,  as.integer(Inches(4)))
    expect_equal(pic$height, as.integer(Inches(2)))
  })

  it("round-trips through save/reload", {
    skip_if_not_installed("magick")
    f   <- make_png()
    prs <- pptx_presentation()
    sl  <- invisible(prs$slides$add_slide(prs$slide_layouts[[6]]))
    invisible(sl$shapes$add_picture(f, Inches(1), Inches(1)))
    tmp <- tempfile(fileext = ".pptx")
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    shapes <- prs2$slides[[1]]$shapes$to_list()
    pics <- Filter(function(s) inherits(s, "Picture"), shapes)
    expect_equal(length(pics), 1L)
  })
})


describe("FreeformBuilder", {
  it("creates a shape via build_freeform + add_line_segments + convert_to_shape", {
    prs   <- pptx_presentation()
    slide <- prs$slides$add_slide(prs$slide_layouts[[6]])
    scale <- Inches(1) / 100
    ff    <- slide$shapes$build_freeform(0, 0, scale)
    expect_s3_class(ff, "FreeformBuilder")
    ff$add_line_segments(list(c(100, 0), c(50, 87), c(0, 0)))
    shp <- ff$convert_to_shape(Inches(1), Inches(1))
    expect_s3_class(shp, "Shape")
    expect_match(shp$name, "Freeform")
    expect_equal(shp$left, as.integer(Inches(1)))
  })

  it("round-trips through save/reload with custom geometry", {
    prs   <- pptx_presentation()
    slide <- prs$slides$add_slide(prs$slide_layouts[[6]])
    ff    <- slide$shapes$build_freeform(0, 0, Inches(1) / 100)
    ff$add_line_segments(list(c(100, 0), c(50, 87), c(0, 0)))
    invisible(ff$convert_to_shape(Inches(1), Inches(1)))
    tmp <- tempfile(fileext = ".pptx")
    prs$save(tmp)
    prs2   <- pptx_presentation(tmp)
    shapes <- prs2$slides[[1]]$shapes$to_list()
    # Find the freeform shape
    ff_shp <- Filter(function(s) grepl("Freeform", s$name), shapes)
    expect_equal(length(ff_shp), 1L)
    elm <- ff_shp[[1]]$.__enclos_env__$private$.element
    expect_true(elm$has_custom_geometry)
  })

  it("returns self from add_line_segments and move_to for chaining", {
    prs   <- pptx_presentation()
    slide <- prs$slides$add_slide(prs$slide_layouts[[6]])
    ff    <- slide$shapes$build_freeform()
    ret   <- ff$add_line_segments(list(c(100, 0), c(0, 100)), close = FALSE)
    expect_identical(ret, ff)
    ret2  <- ff$move_to(50, 50)
    expect_identical(ret2, ff)
  })

  it("shape_offset_x and shape_offset_y return min extents", {
    prs   <- pptx_presentation()
    slide <- prs$slides$add_slide(prs$slide_layouts[[6]])
    ff    <- slide$shapes$build_freeform(100, 200)
    ff$add_line_segments(list(c(300, 400), c(50, 600)), close = FALSE)
    expect_equal(as.integer(ff$shape_offset_x()), 50L)
    expect_equal(as.integer(ff$shape_offset_y()), 200L)
  })
})
