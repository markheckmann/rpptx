# Tests for table support.
# Covers: CT_Table*, SlideShapes$add_table(), Table, TableRow, TableColumn, TableCell.

pptx_path <- function(name) system.file("test_files", name, package = "rpptx")

# Helper: fresh blank presentation with one slide
new_blank_slide <- function() {
  prs    <- pptx_presentation(pptx_path("no-slides.pptx"))
  layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
  prs$slides$add_slide(layout)
}


# ============================================================================
# CT_Table_new_tbl() — XML factory
# ============================================================================

describe("CT_Table_new_tbl()", {
  tbl <- CT_Table_new_tbl(3L, 4L, Inches(8), Inches(4))

  it("returns a CT_Table", {
    expect_true(inherits(tbl, "CT_Table"))
  })

  it("has correct number of rows", {
    expect_equal(length(tbl$tr_lst), 3L)
  })

  it("each row has correct number of cells", {
    for (tr in tbl$tr_lst) {
      expect_equal(length(tr$tc_lst), 4L)
    }
  })

  it("tblGrid has correct number of columns", {
    expect_equal(length(tbl$gridCol_lst), 4L)
  })

  it("column widths sum to total width", {
    widths <- vapply(tbl$gridCol_lst, function(gc) as.integer(gc$w), integer(1))
    expect_equal(sum(widths), as.integer(Inches(8)))
  })

  it("row heights sum to total height", {
    heights <- vapply(tbl$tr_lst, function(tr) as.integer(tr$h), integer(1))
    expect_equal(sum(heights), as.integer(Inches(4)))
  })

  it("has tblPr with firstRow=TRUE and bandRow=TRUE", {
    tblPr <- tbl$tblPr
    expect_false(is.null(tblPr))
    expect_true(isTRUE(tblPr$firstRow))
    expect_true(isTRUE(tblPr$bandRow))
  })

  it("has tableStyleId child in tblPr", {
    tblPr <- tbl$tblPr
    sid <- tblPr$tableStyleId
    expect_false(is.null(sid))
  })

  it("cells are CT_TableCell instances", {
    tc <- tbl$tc(1L, 1L)
    expect_true(inherits(tc, "CT_TableCell"))
  })

  it("each cell has a txBody", {
    tc <- tbl$tc(1L, 1L)
    expect_false(is.null(tc$txBody))
  })
})


describe("CT_TableCell merge attributes", {
  it("gridSpan defaults to 1", {
    tbl <- CT_Table_new_tbl(2L, 2L, Inches(4), Inches(2))
    expect_equal(tbl$tc(1L, 1L)$gridSpan, 1L)
  })

  it("rowSpan defaults to 1", {
    tbl <- CT_Table_new_tbl(2L, 2L, Inches(4), Inches(2))
    expect_equal(tbl$tc(1L, 1L)$rowSpan, 1L)
  })

  it("hMerge defaults to FALSE", {
    tbl <- CT_Table_new_tbl(2L, 2L, Inches(4), Inches(2))
    expect_false(tbl$tc(1L, 1L)$hMerge)
  })

  it("is_merge_origin is TRUE for fresh cell", {
    tbl <- CT_Table_new_tbl(2L, 2L, Inches(4), Inches(2))
    expect_true(tbl$tc(1L, 1L)$is_merge_origin)
  })

  it("is_spanned is FALSE for fresh cell", {
    tbl <- CT_Table_new_tbl(2L, 2L, Inches(4), Inches(2))
    expect_false(tbl$tc(1L, 1L)$is_spanned)
  })
})


# ============================================================================
# SlideShapes$add_table()
# ============================================================================

describe("SlideShapes$add_table()", {
  it("returns a GraphicFrame", {
    slide <- new_blank_slide()
    gf <- slide$shapes$add_table(3L, 4L, Inches(1), Inches(1), Inches(8), Inches(4))
    expect_true(inherits(gf, "GraphicFrame"))
  })

  it("has_table is TRUE", {
    slide <- new_blank_slide()
    gf <- slide$shapes$add_table(3L, 4L, Inches(1), Inches(1), Inches(8), Inches(4))
    expect_true(gf$has_table)
  })

  it("increments slide shape count", {
    slide  <- new_blank_slide()
    before <- length(slide$shapes)
    slide$shapes$add_table(2L, 3L, Inches(1), Inches(1), Inches(6), Inches(3))
    expect_equal(length(slide$shapes), before + 1L)
  })

  it("has correct position", {
    slide <- new_blank_slide()
    gf <- slide$shapes$add_table(2L, 2L, Inches(2), Inches(3), Inches(4), Inches(2))
    expect_equal(as.integer(gf$left),   as.integer(Inches(2)))
    expect_equal(as.integer(gf$top),    as.integer(Inches(3)))
    expect_equal(as.integer(gf$width),  as.integer(Inches(4)))
    expect_equal(as.integer(gf$height), as.integer(Inches(2)))
  })

  it("shape_type is TABLE", {
    slide <- new_blank_slide()
    gf <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    expect_equal(gf$shape_type, MSO_SHAPE_TYPE$TABLE)
  })

  it("name follows 'Table N' pattern", {
    slide <- new_blank_slide()
    gf <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    expect_match(gf$name, "^Table \\d+$")
  })
})


# ============================================================================
# Table — domain object
# ============================================================================

describe("GraphicFrame$table", {
  it("returns a Table object", {
    slide <- new_blank_slide()
    gf <- slide$shapes$add_table(2L, 3L, Inches(1), Inches(1), Inches(6), Inches(3))
    tbl <- gf$table
    expect_true(inherits(tbl, "Table"))
  })

  it("errors when graphicFrame has no tbl element", {
    # Create a minimal graphicFrame manually (no a:tbl inside)
    a <- .nsmap[["a"]]; p <- .nsmap[["p"]]
    xml_str <- sprintf(paste0(
      '<p:graphicFrame xmlns:p="%s" xmlns:a="%s">',
        '<p:nvGraphicFramePr>',
          '<p:cNvPr id="99" name="Frame 98"/>',
          '<p:cNvGraphicFramePr/>',
          '<p:nvPr/>',
        '</p:nvGraphicFramePr>',
        '<p:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></p:xfrm>',
        '<a:graphic><a:graphicData uri="x"/></a:graphic>',
      '</p:graphicFrame>'
    ), p, a)
    gf_elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    # Wrap in a GraphicFrame proxy (parent doesn't matter for this test)
    gf <- GraphicFrame$new(gf_elm, NULL)
    expect_error(gf$table, "table")
  })
})

describe("Table$cell()", {
  it("returns a TableCell", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(3L, 4L, Inches(1), Inches(1), Inches(8), Inches(4))
    tbl <- gf$table
    cell <- tbl$cell(1L, 1L)
    expect_true(inherits(cell, "TableCell"))
  })

  it("errors on out-of-range row", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    tbl <- gf$table
    expect_error(tbl$cell(3L, 1L), "row")
  })

  it("errors on out-of-range col", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    tbl <- gf$table
    expect_error(tbl$cell(1L, 3L), "col")
  })
})

describe("Table$rows and Table$columns", {
  setup_tbl <- function() {
    slide <- new_blank_slide()
    gf <- slide$shapes$add_table(3L, 4L, Inches(1), Inches(1), Inches(8), Inches(4))
    gf$table
  }

  it("rows returns TableRows", {
    tbl <- setup_tbl()
    expect_true(inherits(tbl$rows, "TableRows"))
  })

  it("columns returns TableColumns", {
    tbl <- setup_tbl()
    expect_true(inherits(tbl$columns, "TableColumns"))
  })

  it("length(rows) matches row count", {
    tbl <- setup_tbl()
    expect_equal(length(tbl$rows), 3L)
  })

  it("length(columns) matches column count", {
    tbl <- setup_tbl()
    expect_equal(length(tbl$columns), 4L)
  })

  it("rows[[i]] returns a TableRow", {
    tbl <- setup_tbl()
    expect_true(inherits(tbl$rows[[1]], "TableRow"))
  })

  it("columns[[i]] returns a TableColumn", {
    tbl <- setup_tbl()
    expect_true(inherits(tbl$columns[[1]], "TableColumn"))
  })
})

describe("Table style flags", {
  it("first_row defaults TRUE (from default tblPr)", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    expect_true(gf$table$first_row)
  })

  it("first_row is writable", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    tbl <- gf$table
    tbl$first_row <- FALSE
    expect_false(tbl$first_row)
  })

  it("horz_banding defaults TRUE", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    expect_true(gf$table$horz_banding)
  })

  it("vert_banding defaults FALSE", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    expect_false(gf$table$vert_banding)
  })

  it("vert_banding is writable", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    tbl <- gf$table
    tbl$vert_banding <- TRUE
    expect_true(tbl$vert_banding)
  })
})

describe("Table$iter_cells()", {
  it("returns correct number of cells (rows × cols)", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(3L, 4L, Inches(1), Inches(1), Inches(8), Inches(4))
    cells <- gf$table$iter_cells()
    expect_equal(length(cells), 12L)
  })

  it("all items are TableCell instances", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 3L, Inches(1), Inches(1), Inches(6), Inches(3))
    for (cell in gf$table$iter_cells()) {
      expect_true(inherits(cell, "TableCell"))
    }
  })
})


# ============================================================================
# TableCell
# ============================================================================

describe("TableCell$text", {
  it("text is initially empty string", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    expect_equal(gf$table$cell(1L, 1L)$text, "")
  })

  it("text can be set and retrieved", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    cell <- gf$table$cell(1L, 1L)
    cell$text <- "Hello"
    expect_equal(cell$text, "Hello")
  })

  it("different cells can have different text", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    tbl <- gf$table
    # Use intermediate variables to avoid R's write-back issue with method calls
    c11 <- tbl$cell(1L, 1L); c11$text <- "A"
    c12 <- tbl$cell(1L, 2L); c12$text <- "B"
    c21 <- tbl$cell(2L, 1L); c21$text <- "C"
    expect_equal(tbl$cell(1L, 1L)$text, "A")
    expect_equal(tbl$cell(1L, 2L)$text, "B")
    expect_equal(tbl$cell(2L, 1L)$text, "C")
  })
})

describe("TableCell$text_frame", {
  it("returns a TextFrame", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    tf <- gf$table$cell(1L, 1L)$text_frame
    expect_true(inherits(tf, "TextFrame"))
  })

  it("text_frame text is accessible and writable", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    cell <- gf$table$cell(1L, 1L)
    cell$text_frame$text <- "World"
    expect_equal(cell$text, "World")
  })
})

describe("TableCell$fill", {
  it("fill returns a FillFormat", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    ff <- gf$table$cell(1L, 1L)$fill
    expect_true(inherits(ff, "FillFormat"))
  })

  it("cell fill can be set to solid with a color", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    cell <- gf$table$cell(1L, 1L)
    cell$fill$solid()
    cell$fill$fore_color$rgb <- RGBColor(255L, 0L, 0L)
    expect_equal(cell$fill$type, MSO_FILL$SOLID)
    expect_true(cell$fill$fore_color$rgb == RGBColor(255L, 0L, 0L))
  })
})

describe("TableCell$is_merge_origin / is_spanned", {
  it("fresh cells are merge-origin", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    expect_true(gf$table$cell(1L, 1L)$is_merge_origin)
  })

  it("fresh cells are not spanned", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    expect_false(gf$table$cell(1L, 1L)$is_spanned)
  })
})


# ============================================================================
# TableRow / TableColumn
# ============================================================================

describe("TableRow$height", {
  it("height is an Emu value", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(3L, 2L, Inches(1), Inches(1), Inches(4), Inches(3))
    h <- gf$table$rows[[1]]$height
    expect_true(is.numeric(h) || inherits(h, "Emu"))
  })

  it("height can be set and retrieved", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(3L, 2L, Inches(1), Inches(1), Inches(4), Inches(3))
    row <- gf$table$rows[[2]]
    row$height <- Inches(1)
    expect_equal(as.integer(row$height), as.integer(Inches(1)))
  })

  it("row cells are accessible", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 3L, Inches(1), Inches(1), Inches(6), Inches(2))
    cells <- gf$table$rows[[1]]$cells
    expect_true(inherits(cells, "TableCells"))
    expect_equal(length(cells), 3L)
  })
})

describe("TableColumn$width", {
  it("width is an Emu value", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 4L, Inches(1), Inches(1), Inches(8), Inches(2))
    w <- gf$table$columns[[1]]$width
    expect_true(is.numeric(w) || inherits(w, "Emu"))
  })

  it("width can be set and retrieved", {
    slide <- new_blank_slide()
    gf  <- slide$shapes$add_table(2L, 4L, Inches(1), Inches(1), Inches(8), Inches(2))
    col <- gf$table$columns[[2]]
    col$width <- Inches(3)
    expect_equal(as.integer(col$width), as.integer(Inches(3)))
  })
})


# ============================================================================
# Round-trip: save and reload
# ============================================================================

describe("Table round-trip (save/reload)", {
  it("table text is preserved after save/reload", {
    withr::with_tempfile("tmp", fileext = ".pptx", {
      prs    <- pptx_presentation(pptx_path("no-slides.pptx"))
      layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
      slide  <- prs$slides$add_slide(layout)
      gf     <- slide$shapes$add_table(2L, 2L,
                                        Inches(1), Inches(1), Inches(4), Inches(2))
      tbl <- gf$table
      c11 <- tbl$cell(1L, 1L); c11$text <- "Header"
      c21 <- tbl$cell(2L, 1L); c21$text <- "Data"
      prs$save(tmp)

      prs2   <- pptx_presentation(tmp)
      slide2 <- prs2$slides[[1]]
      gf2    <- Filter(function(s) isTRUE(s$has_table), slide2$shapes$to_list())[[1]]
      tbl2   <- gf2$table

      cell_11 <- tbl2$cell(1L, 1L)
      cell_21 <- tbl2$cell(2L, 1L)
      expect_equal(cell_11$text, "Header")
      expect_equal(cell_21$text, "Data")
    })
  })
})


# ============================================================================
# Table$style_id
# ============================================================================

describe("Table$style_id", {
  it("returns a non-NULL GUID for a newly created table", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    gf     <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    sid    <- gf$table$style_id
    expect_type(sid, "character")
    expect_true(grepl("\\{", sid))
  })

  it("style_id can be changed", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    gf     <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    new_id <- "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    gf$table$style_id <- new_id
    expect_equal(gf$table$style_id, new_id)
  })

  it("style_id round-trips through save/load", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    gf     <- slide$shapes$add_table(2L, 2L, Inches(1), Inches(1), Inches(4), Inches(2))
    custom_id <- "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    gf$table$style_id <- custom_id
    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2   <- pptx_presentation(tmp)
    gf2    <- Filter(function(s) isTRUE(s$has_table), prs2$slides[[1]]$shapes$to_list())[[1]]
    expect_equal(gf2$table$style_id, custom_id)
  })
})
