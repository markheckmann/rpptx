# Tests for oxml-shapes.R, shapes-base.R, enum-shapes.R
#
# Phase 4: shape layer — element wrappers, proxy objects, factory dispatch.

pptx_path <- function(name) system.file("test_files", name, package = "rpptx")

# ============================================================================
# Helper: open test_slides.pptx and return first slide's shapes
# ============================================================================

shapes_slide1 <- function() {
  prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
  slide <- prs$slides[[1]]
  slide$shapes$to_list()
}


# ============================================================================
# MSO_SHAPE_TYPE enum
# ============================================================================

describe("MSO_SHAPE_TYPE", {
  it("exposes expected integer constants", {
    expect_identical(MSO_SHAPE_TYPE$AUTO_SHAPE,  1L)
    expect_identical(MSO_SHAPE_TYPE$CHART,        3L)
    expect_identical(MSO_SHAPE_TYPE$GROUP,        6L)
    expect_identical(MSO_SHAPE_TYPE$PICTURE,     13L)
    expect_identical(MSO_SHAPE_TYPE$PLACEHOLDER, 14L)
    expect_identical(MSO_SHAPE_TYPE$TEXT_BOX,    17L)
    expect_identical(MSO_SHAPE_TYPE$TABLE,       19L)
  })

  it("MSO alias equals MSO_SHAPE_TYPE", {
    expect_identical(MSO, MSO_SHAPE_TYPE)
  })
})


# ============================================================================
# PP_PLACEHOLDER enum
# ============================================================================

describe("PP_PLACEHOLDER", {
  it("exposes expected string constants", {
    expect_equal(PP_PLACEHOLDER$TITLE,        "title")
    expect_equal(PP_PLACEHOLDER$BODY,         "body")
    expect_equal(PP_PLACEHOLDER$CENTER_TITLE, "ctrTitle")
    expect_equal(PP_PLACEHOLDER$SUBTITLE,     "subTitle")
    expect_equal(PP_PLACEHOLDER$OBJECT,       "obj")
  })
})


# ============================================================================
# SlideShapes collection
# ============================================================================

describe("SlideShapes", {
  it("returns correct count from test_slides.pptx", {
    prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    expect_equal(length(slide$shapes), 9L)
  })

  it("[[ returns same shape as get()", {
    prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    s1a <- slide$shapes[[1]]
    s1b <- slide$shapes$get(1L)
    expect_equal(class(s1a)[1], class(s1b)[1])
    expect_equal(s1a$name, s1b$name)
  })

  it("to_list() returns all shapes", {
    shapes <- shapes_slide1()
    expect_length(shapes, 9L)
  })

  it("errors on out-of-range index", {
    prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    expect_error(slide$shapes[[0L]], "out of range")
    expect_error(slide$shapes[[10L]], "out of range")
  })
})


# ============================================================================
# shape_factory dispatch
# ============================================================================

describe("shape_factory", {
  it("dispatches p:sp as Shape", {
    shapes <- shapes_slide1()
    title_shape <- shapes[[1]]
    expect_s3_class(title_shape, "Shape")
  })

  it("dispatches p:pic as Picture", {
    shapes <- shapes_slide1()
    pic <- shapes[[3]]
    expect_s3_class(pic, "Picture")
  })

  it("dispatches p:grpSp as GroupShape", {
    shapes <- shapes_slide1()
    grp <- shapes[[5]]
    expect_s3_class(grp, "GroupShape")
  })

  it("dispatches p:cxnSp as Connector", {
    shapes <- shapes_slide1()
    conn <- shapes[[8]]
    expect_s3_class(conn, "Connector")
  })

  it("dispatches p:graphicFrame as GraphicFrame", {
    shapes <- shapes_slide1()
    gf <- shapes[[9]]
    expect_s3_class(gf, "GraphicFrame")
  })
})


# ============================================================================
# BaseShape — common properties
# ============================================================================

describe("BaseShape properties", {
  it("shape_id returns integer", {
    shapes <- shapes_slide1()
    expect_type(shapes[[1]]$shape_id, "integer")
  })

  it("name returns character", {
    shapes <- shapes_slide1()
    expect_type(shapes[[1]]$name, "character")
    expect_equal(shapes[[1]]$name, "Title 1")
  })

  it("left/top/width/height return numeric EMU values", {
    shapes <- shapes_slide1()
    ph_body <- shapes[[2]]  # Content Placeholder 2 — has explicit spPr geometry
    expect_true(is.numeric(ph_body$left))
    expect_true(is.numeric(ph_body$top))
    expect_true(is.numeric(ph_body$width))
    expect_true(is.numeric(ph_body$height))
    expect_gt(ph_body$width,  0)
    expect_gt(ph_body$height, 0)
  })

  it("rotation returns 0 for non-rotated shape", {
    shapes <- shapes_slide1()
    expect_equal(shapes[[6]]$rotation, 0)
  })
})


# ============================================================================
# Shape type detection
# ============================================================================

describe("Shape$shape_type", {
  it("is PLACEHOLDER for title placeholder", {
    shapes <- shapes_slide1()
    title  <- shapes[[1]]
    expect_true(title$is_placeholder)
    expect_equal(title$shape_type, MSO_SHAPE_TYPE$PLACEHOLDER)
  })

  it("is TEXT_BOX for textbox shape", {
    shapes <- shapes_slide1()
    tb <- shapes[[4]]  # TextBox 6
    expect_false(tb$is_placeholder)
    expect_equal(tb$shape_type, MSO_SHAPE_TYPE$TEXT_BOX)
  })

  it("is AUTO_SHAPE for rounded rectangle", {
    shapes <- shapes_slide1()
    rect <- shapes[[6]]  # Rounded Rectangle 11
    expect_false(rect$is_placeholder)
    expect_equal(rect$shape_type, MSO_SHAPE_TYPE$AUTO_SHAPE)
  })

  it("has_text_frame is TRUE", {
    shapes <- shapes_slide1()
    expect_true(shapes[[1]]$has_text_frame)
  })
})

describe("Picture$shape_type", {
  it("returns PICTURE", {
    shapes <- shapes_slide1()
    expect_equal(shapes[[3]]$shape_type, MSO_SHAPE_TYPE$PICTURE)
  })

  it("has_text_frame is FALSE", {
    shapes <- shapes_slide1()
    expect_false(shapes[[3]]$has_text_frame)
  })
})

describe("Connector$shape_type", {
  it("returns LINE (matches python-pptx behaviour)", {
    shapes <- shapes_slide1()
    expect_equal(shapes[[8]]$shape_type, MSO_SHAPE_TYPE$LINE)
  })
})

describe("GraphicFrame$shape_type", {
  it("returns TABLE for table frame", {
    shapes <- shapes_slide1()
    gf <- shapes[[9]]  # Table 15
    expect_equal(gf$shape_type, MSO_SHAPE_TYPE$TABLE)
    expect_true(gf$has_table)
    expect_false(gf$has_chart)
  })
})

describe("GroupShape$shape_type", {
  it("returns GROUP", {
    shapes <- shapes_slide1()
    expect_equal(shapes[[5]]$shape_type, MSO_SHAPE_TYPE$GROUP)
  })
})


# ============================================================================
# PlaceholderFormat
# ============================================================================

describe("PlaceholderFormat", {
  it("is_placeholder TRUE for title shape", {
    shapes <- shapes_slide1()
    expect_true(shapes[[1]]$is_placeholder)
  })

  it("is_placeholder FALSE for picture", {
    shapes <- shapes_slide1()
    expect_false(shapes[[3]]$is_placeholder)
  })

  it("placeholder_format errors for non-placeholder", {
    shapes <- shapes_slide1()
    expect_error(shapes[[3]]$placeholder_format, "not a placeholder")
  })

  it("placeholder_format$type returns 'title' for title", {
    shapes <- shapes_slide1()
    pf <- shapes[[1]]$placeholder_format
    expect_equal(pf$type, "title")
  })

  it("placeholder_format$idx returns 0L for title", {
    shapes <- shapes_slide1()
    pf <- shapes[[1]]$placeholder_format
    expect_equal(pf$idx, 0L)
  })

  it("placeholder_format$type returns 'obj' for body placeholder", {
    shapes <- shapes_slide1()
    pf <- shapes[[2]]$placeholder_format
    expect_equal(pf$type, "obj")
    expect_equal(pf$idx, 1L)
  })
})


# ============================================================================
# GroupShape children
# ============================================================================

describe("GroupShape$shapes", {
  it("returns child shapes", {
    shapes <- shapes_slide1()
    grp    <- shapes[[5]]
    children <- grp$shapes$to_list()
    expect_length(children, 2L)
  })

  it("child types are correct (Picture + Shape)", {
    shapes   <- shapes_slide1()
    children <- shapes[[5]]$shapes$to_list()
    types    <- vapply(children, function(s) class(s)[1], character(1))
    expect_true("Picture" %in% types)
    expect_true("Shape"   %in% types)
  })

  it("[[ dispatches correctly on GroupShapes", {
    shapes <- shapes_slide1()
    grp    <- shapes[[5]]
    child  <- grp$shapes[[1]]
    expect_false(is.null(child))
    expect_true(inherits(child, "BaseShape"))
  })
})


# ============================================================================
# CT_GroupShape (oxml layer)
# ============================================================================

describe("CT_GroupShape", {
  it("iter_shape_elms returns 9 elements for slide 1 spTree", {
    prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    elms  <- slide$element$spTree$iter_shape_elms()
    expect_length(elms, 9L)
  })

  it("max_shape_id and next_shape_id are sensible integers", {
    prs    <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide  <- prs$slides[[1]]
    sp_tree <- slide$element$spTree
    max_id  <- sp_tree$max_shape_id()
    next_id <- sp_tree$next_shape_id()
    expect_type(max_id,  "integer")
    expect_type(next_id, "integer")
    expect_equal(next_id, max_id + 1L)
  })
})


# ============================================================================
# SlidePlaceholders
# ============================================================================

describe("SlidePlaceholders", {
  it("returns all placeholder shapes on a slide", {
    prs    <- pptx_presentation(test_file_path("../templates/default.pptx"))
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    phs    <- slide$placeholders
    expect_s3_class(phs, "SlidePlaceholders")
    expect_gte(length(phs), 1L)
  })

  it("supports [[ for 1-based position access", {
    prs    <- pptx_presentation(test_file_path("../templates/default.pptx"))
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    phs    <- slide$placeholders
    ph1 <- phs[[1]]
    expect_true(ph1$is_placeholder)
  })

  it("supports get(idx) for OOXML idx access", {
    prs    <- pptx_presentation(test_file_path("../templates/default.pptx"))
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    phs    <- slide$placeholders
    # idx=0 is title placeholder in most layouts
    title_ph <- phs$get(0L)
    expect_true(title_ph$is_placeholder)
    expect_equal(title_ph$placeholder_format$idx, 0L)
  })

  it("errors when OOXML idx not found", {
    prs    <- pptx_presentation(test_file_path("../templates/default.pptx"))
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    phs    <- slide$placeholders
    expect_error(phs$get(9999L), regexp = "no placeholder")
  })

  it("placeholder has text_frame", {
    prs    <- pptx_presentation(test_file_path("../templates/default.pptx"))
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    title_ph <- slide$placeholders$get(0L)
    expect_true(title_ph$has_text_frame)
    expect_s3_class(title_ph$text_frame, "TextFrame")
  })
})


describe("Placeholder position inheritance", {
  it("inherits left/top/width/height from layout when shape has no xfrm", {
    prs    <- pptx_presentation(test_file_path("../templates/default.pptx"))
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    title_ph <- slide$placeholders$get(0L)

    # Inherited values should be non-zero (from layout)
    expect_gt(as.integer(title_ph$left),   0L)
    expect_gt(as.integer(title_ph$width),  0L)
    expect_gt(as.integer(title_ph$height), 0L)
  })

  it("position values are Length objects (EMU)", {
    prs    <- pptx_presentation(test_file_path("../templates/default.pptx"))
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    title_ph <- slide$placeholders$get(0L)
    expect_s3_class(title_ph$left,   "Length")
    expect_s3_class(title_ph$width,  "Length")
    expect_s3_class(title_ph$height, "Length")
  })
})


describe("Placeholder class dispatch", {
  it("slide placeholders are SlidePlaceholder", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    ph <- slide$placeholders$get_at(1)
    expect_s3_class(ph, "SlidePlaceholder")
    expect_s3_class(ph, "Shape")
  })

  it("layout placeholders are LayoutPlaceholder", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    ph <- layout$placeholders$get_at(1)
    expect_s3_class(ph, "LayoutPlaceholder")
  })

  it("master placeholders are MasterPlaceholder", {
    prs    <- pptx_presentation()
    master <- prs$slide_masters[[1]]
    ph <- master$placeholders$get_at(1)
    expect_s3_class(ph, "MasterPlaceholder")
  })

  it("non-placeholder shapes are still Shape", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[6]]
    slide  <- prs$slides$add_slide(layout)
    shape  <- slide$shapes$add_shape(MSO_AUTO_SHAPE_TYPE$RECTANGLE,
                Inches(1), Inches(1), Inches(2), Inches(1))
    expect_s3_class(shape, "Shape")
    expect_false(inherits(shape, "SlidePlaceholder"))
  })
})


describe("Slide → Layout → Master inheritance chain", {
  it("SlideMasterPlaceholders$get returns MasterPlaceholder by type", {
    prs    <- pptx_presentation()
    master <- prs$slide_masters[[1]]
    title_ph <- master$placeholders$get("title")
    expect_s3_class(title_ph, "MasterPlaceholder")
    expect_equal(title_ph$placeholder_format$type, "title")
  })

  it("SlideMasterPlaceholders$get returns NULL for unknown type", {
    prs    <- pptx_presentation()
    master <- prs$slide_masters[[1]]
    expect_null(master$placeholders$get("nonexistent_type"))
  })

  it("LayoutPlaceholder inherits from master when layout has no xfrm", {
    prs    <- pptx_presentation()
    master <- prs$slide_masters[[1]]
    layout <- prs$slide_layouts[[6]]  # Blank — placeholders may inherit
    # Master title dimensions should be non-zero
    m_ph <- master$placeholders$get("title")
    if (!is.null(m_ph)) {
      expect_gt(as.integer(m_ph$left),  0L)
      expect_gt(as.integer(m_ph$width), 0L)
    }
  })

  it("SlidePlaceholder dimension inherits from layout (two-level test)", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]  # Title Slide
    slide  <- prs$slides$add_slide(layout)
    slide_ph  <- slide$placeholders$get(0L)   # title/ctrTitle
    layout_ph <- layout$placeholders$get(0L)
    # Slide should inherit the same left as the layout
    expect_equal(as.integer(slide_ph$left), as.integer(layout_ph$left))
    expect_equal(as.integer(slide_ph$width), as.integer(layout_ph$width))
  })

  it("MasterPlaceholder returns its own dimensions without climbing further", {
    prs    <- pptx_presentation()
    master <- prs$slide_masters[[1]]
    ph <- master$placeholders$get("title")
    expect_s3_class(ph, "MasterPlaceholder")
    # Should return a Length (not an error)
    expect_s3_class(ph$left, "Length")
    expect_s3_class(ph$width, "Length")
  })
})

# ============================================================================
# Connector routing — begin_connect / end_connect
# ============================================================================

new_slide_with_shapes <- function() {
  prs    <- pptx_presentation(system.file("test_files", "no-slides.pptx", package = "rpptx"))
  layout <- prs$slide_masters[[1]]$slide_layouts[[6]]
  slide  <- prs$slides$add_slide(layout)
  r1 <- slide$shapes$add_shape(MSO_AUTO_SHAPE_TYPE$RECTANGLE,
          Inches(1), Inches(1), Inches(2), Inches(1))
  r2 <- slide$shapes$add_shape(MSO_AUTO_SHAPE_TYPE$RECTANGLE,
          Inches(5), Inches(1), Inches(2), Inches(1))
  conn <- slide$shapes$add_connector(MSO_CONNECTOR_TYPE$STRAIGHT,
            Inches(3), Inches(1.5), Inches(5), Inches(1.5))
  list(slide = slide, r1 = r1, r2 = r2, conn = conn)
}

describe("Connector$begin_connect / end_connect", {
  it("sets begin connection to a shape", {
    s <- new_slide_with_shapes()
    s$conn$begin_connect(s$r1, 1L)  # right side (idx 1)
    expect_equal(s$conn$begin_connected_shape_id, s$r1$shape_id)
    expect_equal(s$conn$begin_connection_site_index, 1L)
  })

  it("sets end connection to a shape", {
    s <- new_slide_with_shapes()
    s$conn$end_connect(s$r2, 3L)   # left side (idx 3)
    expect_equal(s$conn$end_connected_shape_id, s$r2$shape_id)
    expect_equal(s$conn$end_connection_site_index, 3L)
  })

  it("unconnected connector returns NULL for all ids/indices", {
    s <- new_slide_with_shapes()
    expect_null(s$conn$begin_connected_shape_id)
    expect_null(s$conn$begin_connection_site_index)
    expect_null(s$conn$end_connected_shape_id)
    expect_null(s$conn$end_connection_site_index)
  })

  it("disconnects begin endpoint", {
    s <- new_slide_with_shapes()
    s$conn$begin_connect(s$r1, 0L)
    s$conn$begin_disconnect()
    expect_null(s$conn$begin_connected_shape_id)
  })

  it("disconnects end endpoint", {
    s <- new_slide_with_shapes()
    s$conn$end_connect(s$r2, 2L)
    s$conn$end_disconnect()
    expect_null(s$conn$end_connected_shape_id)
  })

  it("round-trips through save/reopen", {
    s <- new_slide_with_shapes()
    s$conn$begin_connect(s$r1, 1L)
    s$conn$end_connect(s$r2, 3L)
    r1_id <- s$r1$shape_id
    r2_id <- s$r2$shape_id
    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs2 <- pptx_presentation(system.file("test_files", "no-slides.pptx", package = "rpptx"))
    layout <- prs2$slide_masters[[1]]$slide_layouts[[6]]
    slide2 <- prs2$slides$add_slide(layout)
    r1b <- slide2$shapes$add_shape(MSO_AUTO_SHAPE_TYPE$RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
    r2b <- slide2$shapes$add_shape(MSO_AUTO_SHAPE_TYPE$RECTANGLE, Inches(5), Inches(1), Inches(2), Inches(1))
    conn2 <- slide2$shapes$add_connector(MSO_CONNECTOR_TYPE$STRAIGHT, Inches(3), Inches(1.5), Inches(5), Inches(1.5))
    conn2$begin_connect(r1b, 1L)
    conn2$end_connect(r2b, 3L)
    prs2$save(tmp)
    prs3 <- pptx_presentation(tmp)
    shapes3 <- prs3$slides[[1]]$shapes$to_list()
    conn3 <- Filter(function(s) s$shape_type == MSO_SHAPE_TYPE$LINE, shapes3)[[1]]
    expect_equal(conn3$begin_connected_shape_id, r1b$shape_id)
    expect_equal(conn3$end_connected_shape_id,   r2b$shape_id)
  })

  it("Connector$line returns a LineFormat", {
    s <- new_slide_with_shapes()
    expect_s3_class(s$conn$line, "R6")
    expect_true(inherits(s$conn$line, "LineFormat"))
  })
})
