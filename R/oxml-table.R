# Custom XML element classes for table elements.
#
# Ported from python-pptx/src/pptx/oxml/table.py.

# ============================================================================
# CT_TableProperties — <a:tblPr>
# ============================================================================

#' @include oxml-xmlchemy.R oxml-simpletypes.R
#' @noRd
CT_TableProperties <- define_oxml_element(
  classname  = "CT_TableProperties",
  tag        = "a:tblPr",
  children   = list(
    tableStyleId = zero_or_one("a:tableStyleId", successors = c("a:extLst"))
  ),
  attributes = list(
    firstRow = optional_attribute("firstRow", XsdBoolean, default = FALSE),
    lastRow  = optional_attribute("lastRow",  XsdBoolean, default = FALSE),
    firstCol = optional_attribute("firstCol", XsdBoolean, default = FALSE),
    lastCol  = optional_attribute("lastCol",  XsdBoolean, default = FALSE),
    bandRow  = optional_attribute("bandRow",  XsdBoolean, default = FALSE),
    bandCol  = optional_attribute("bandCol",  XsdBoolean, default = FALSE)
  )
)


# ============================================================================
# CT_TableGrid — <a:tblGrid>
# ============================================================================

#' @noRd
CT_TableGrid <- define_oxml_element(
  classname = "CT_TableGrid",
  tag       = "a:tblGrid",
  children  = list(
    gridCol = zero_or_more("a:gridCol")
  )
)

# Add a gridCol with the specified width in EMU
CT_TableGrid$set("public", "add_gridCol", function(width) {
  xmlns_a <- .nsmap[["a"]]
  xml2::xml_add_child(
    private$.node,
    xml2::read_xml(sprintf('<a:gridCol xmlns:a="%s" w="%d"/>', xmlns_a, as.integer(width)))
  )
  wrap_element(xml2::xml_child(private$.node, xml2::xml_length(private$.node)))
})


# ============================================================================
# CT_TableCol — <a:gridCol>
# ============================================================================

#' @noRd
CT_TableCol <- define_oxml_element(
  classname  = "CT_TableCol",
  tag        = "a:gridCol",
  attributes = list(
    w = required_attribute("w", ST_PositiveCoordinate)
  )
)


# ============================================================================
# CT_TableCellProperties — <a:tcPr>
# ============================================================================

#' @noRd
CT_TableCellProperties <- define_oxml_element(
  classname  = "CT_TableCellProperties",
  tag        = "a:tcPr",
  attributes = list(
    marL   = optional_attribute("marL",   ST_Coordinate32, default = NULL),
    marR   = optional_attribute("marR",   ST_Coordinate32, default = NULL),
    marT   = optional_attribute("marT",   ST_Coordinate32, default = NULL),
    marB   = optional_attribute("marB",   ST_Coordinate32, default = NULL),
    anchor = optional_attribute("anchor", XsdString,       default = NULL)
  )
)


# ============================================================================
# CT_TableCell — <a:tc>
# ============================================================================

#' CT_TableCell XML element
#' @noRd
CT_TableCell <- R6::R6Class(
  "CT_TableCell",
  inherit = BaseOxmlElement,

  public = list(
    # Return or create <a:txBody>
    get_or_add_txBody = function() {
      existing <- self$txBody
      if (!is.null(existing)) return(existing)
      a <- .nsmap[["a"]]
      node <- xml2::read_xml(sprintf(
        '<a:txBody xmlns:a="%s"><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>', a
      ))
      # Insert before tcPr (or extLst), or prepend
      children <- xml2::xml_children(private$.node)
      inserted <- FALSE
      after_tags <- c(qn("a:tcPr"), qn("a:extLst"))
      for (child in children) {
        if (.get_clark_name(child) %in% after_tags) {
          xml2::xml_add_sibling(child, xml2::xml_root(node), .where = "before")
          inserted <- TRUE
          break
        }
      }
      if (!inserted) {
        xml2::xml_add_child(private$.node, xml2::xml_root(node))
      }
      self$txBody
    },

    # Return or create <a:tcPr>
    get_or_add_tcPr = function() {
      existing <- self$tcPr
      if (!is.null(existing)) return(existing)
      a <- .nsmap[["a"]]
      node <- xml2::read_xml(sprintf('<a:tcPr xmlns:a="%s"/>', a))
      # Insert before extLst, or at end
      children <- xml2::xml_children(private$.node)
      inserted <- FALSE
      for (child in children) {
        if (.get_clark_name(child) == qn("a:extLst")) {
          xml2::xml_add_sibling(child, xml2::xml_root(node), .where = "before")
          inserted <- TRUE
          break
        }
      }
      if (!inserted) {
        xml2::xml_add_child(private$.node, xml2::xml_root(node))
      }
      self$tcPr
    }
  ),

  active = list(
    # <a:txBody> child, or NULL
    txBody = function() {
      r <- xml2::xml_find_first(private$.node, "a:txBody", ns = c(a = .nsmap[["a"]]))
      if (inherits(r, "xml_missing")) return(NULL)
      wrap_element(r)
    },

    # <a:tcPr> child, or NULL
    tcPr = function() {
      r <- xml2::xml_find_first(private$.node, "a:tcPr", ns = c(a = .nsmap[["a"]]))
      if (inherits(r, "xml_missing")) return(NULL)
      wrap_element(r)
    },

    # Number of columns this cell spans (1 = no merge)
    gridSpan = function(value) {
      if (!missing(value)) {
        v <- as.integer(value)
        if (v <= 1L) self$remove_attr("gridSpan") else self$set_attr("gridSpan", as.character(v))
        return(invisible(value))
      }
      v <- self$get_attr("gridSpan")
      if (is.null(v)) 1L else as.integer(v)
    },

    # Number of rows this cell spans (1 = no merge)
    rowSpan = function(value) {
      if (!missing(value)) {
        v <- as.integer(value)
        if (v <= 1L) self$remove_attr("rowSpan") else self$set_attr("rowSpan", as.character(v))
        return(invisible(value))
      }
      v <- self$get_attr("rowSpan")
      if (is.null(v)) 1L else as.integer(v)
    },

    # TRUE if this cell is a non-origin part of a horizontal merge
    hMerge = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) self$set_attr("hMerge", "1") else self$remove_attr("hMerge")
        return(invisible(value))
      }
      v <- self$get_attr("hMerge")
      !is.null(v) && v %in% c("1", "true")
    },

    # TRUE if this cell is a non-origin part of a vertical merge
    vMerge = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) self$set_attr("vMerge", "1") else self$remove_attr("vMerge")
        return(invisible(value))
      }
      v <- self$get_attr("vMerge")
      !is.null(v) && v %in% c("1", "true")
    },

    # TRUE if this is the top-left cell of a merged range (not spanned)
    is_merge_origin = function() !isTRUE(self$hMerge) && !isTRUE(self$vMerge),

    # TRUE if this cell is part of a merge but not the origin
    is_spanned = function() isTRUE(self$hMerge) || isTRUE(self$vMerge)
  )
)


# ============================================================================
# CT_TableRow — <a:tr>
# ============================================================================

#' CT_TableRow XML element
#' @noRd
CT_TableRow <- R6::R6Class(
  "CT_TableRow",
  inherit = BaseOxmlElement,

  public = list(
    # Append a new minimal <a:tc> and return it wrapped
    add_tc = function() {
      a <- .nsmap[["a"]]
      xml_str <- sprintf(
        '<a:tc xmlns:a="%s"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody></a:tc>', a
      )
      xml2::xml_add_child(private$.node, xml2::read_xml(xml_str))
      wrap_element(xml2::xml_child(private$.node, xml2::xml_length(private$.node)))
    }
  ),

  active = list(
    # Row height in EMU (read/write)
    h = function(value) {
      if (!missing(value)) {
        self$set_attr("h", ST_PositiveCoordinate$to_xml(value))
        return(invisible(value))
      }
      h_str <- self$get_attr("h")
      if (is.null(h_str)) NULL else ST_PositiveCoordinate$from_xml(h_str)
    },

    # List of <a:tc> children (wrapped)
    tc_lst = function() {
      r <- xml2::xml_find_all(private$.node, "a:tc", ns = c(a = .nsmap[["a"]]))
      lapply(r, wrap_element)
    }
  )
)


# ============================================================================
# CT_Table — <a:tbl>
# ============================================================================

#' CT_Table XML element
#' @noRd
CT_Table <- R6::R6Class(
  "CT_Table",
  inherit = BaseOxmlElement,

  public = list(
    # Return CT_TableCell at 1-based (row_idx, col_idx)
    tc = function(row_idx, col_idx) {
      rows <- self$tr_lst
      if (row_idx < 1L || row_idx > length(rows)) {
        stop("row index out of range", call. = FALSE)
      }
      tcs <- rows[[row_idx]]$tc_lst
      if (col_idx < 1L || col_idx > length(tcs)) {
        stop("col index out of range", call. = FALSE)
      }
      tcs[[col_idx]]
    },

    # Iterate all cells left-to-right, top-to-bottom
    iter_tcs = function() {
      cells <- list()
      for (tr in self$tr_lst) {
        for (tc in tr$tc_lst) {
          cells <- c(cells, list(tc))
        }
      }
      cells
    }
  ),

  active = list(
    # <a:tblPr> child, or NULL
    tblPr = function() {
      r <- xml2::xml_find_first(private$.node, "a:tblPr", ns = c(a = .nsmap[["a"]]))
      if (inherits(r, "xml_missing")) return(NULL)
      wrap_element(r)
    },

    # <a:tblGrid> child, or NULL
    tblGrid = function() {
      r <- xml2::xml_find_first(private$.node, "a:tblGrid", ns = c(a = .nsmap[["a"]]))
      if (inherits(r, "xml_missing")) return(NULL)
      wrap_element(r)
    },

    # List of <a:tr> children (wrapped)
    tr_lst = function() {
      r <- xml2::xml_find_all(private$.node, "a:tr", ns = c(a = .nsmap[["a"]]))
      lapply(r, wrap_element)
    },

    # List of <a:gridCol> children from tblGrid (wrapped)
    gridCol_lst = function() {
      grid <- self$tblGrid
      if (is.null(grid)) return(list())
      r <- xml2::xml_find_all(grid$get_node(), "a:gridCol", ns = c(a = .nsmap[["a"]]))
      lapply(r, wrap_element)
    }
  )
)


# ============================================================================
# CT_Table factory — build a complete tbl element with rows and columns
# ============================================================================

#' Create a new <a:tbl> element
#' @noRd
CT_Table_new_tbl <- function(rows, cols, width, height) {
  a <- .nsmap[["a"]]
  xml_str <- sprintf(paste0(
    '<a:tbl xmlns:a="%s">',
      '<a:tblPr firstRow="1" bandRow="1">',
        '<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>',
      '</a:tblPr>',
      '<a:tblGrid/>',
    '</a:tbl>'
  ), a)
  tbl <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))

  # Evenly distribute width across columns
  col_w <- as.integer(width) %/% as.integer(cols)
  last_w <- as.integer(width) - col_w * (as.integer(cols) - 1L)
  grid <- tbl$tblGrid
  for (i in seq_len(cols - 1L)) grid$add_gridCol(col_w)
  if (cols >= 1L) grid$add_gridCol(last_w)

  # Evenly distribute height across rows
  row_h <- as.integer(height) %/% as.integer(rows)
  last_h <- as.integer(height) - row_h * (as.integer(rows) - 1L)
  for (i in seq_len(rows)) {
    h <- if (i < rows) row_h else last_h
    tr_str <- sprintf('<a:tr xmlns:a="%s" h="%d"/>', a, as.integer(h))
    xml2::xml_add_child(tbl$get_node(), xml2::read_xml(tr_str))
    tr <- wrap_element(xml2::xml_child(tbl$get_node(), xml2::xml_length(tbl$get_node())))
    for (j in seq_len(cols)) tr$add_tc()
  }

  tbl
}


# ============================================================================
# CT_GraphicalObjectFrame extensions for table support
# ============================================================================

# GRAPHIC_DATA_URI for tables
.GRAPHIC_DATA_URI_TABLE <- "http://schemas.openxmlformats.org/drawingml/2006/table"

# Get or create p:xfrm child on a graphicFrame element
.gf_get_or_add_pxfrm <- function(self_elm) {
  existing <- self_elm$find(qn("p:xfrm"))
  if (!is.null(existing)) return(existing)
  p_ns <- .nsmap[["p"]]; a_ns <- .nsmap[["a"]]
  xfrm_str <- sprintf(
    '<p:xfrm xmlns:p="%s" xmlns:a="%s"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></p:xfrm>',
    p_ns, a_ns
  )
  self_elm$insert_element_before(
    wrap_element(xml2::xml_root(xml2::read_xml(xfrm_str))),
    "a:graphic"
  )
  self_elm$find(qn("p:xfrm"))
}

# Override x/y/cx/cy/rot on CT_GraphicalObjectFrame to use p:xfrm
CT_GraphicalObjectFrame$set("active", "x", function(value) {
  if (!missing(value)) {
    .gf_get_or_add_pxfrm(self)$x <- value; return(invisible(value))
  }
  xfrm <- self$find(qn("p:xfrm")); if (is.null(xfrm)) Emu(0L) else xfrm$x %||% Emu(0L)
}, overwrite = TRUE)

CT_GraphicalObjectFrame$set("active", "y", function(value) {
  if (!missing(value)) {
    .gf_get_or_add_pxfrm(self)$y <- value; return(invisible(value))
  }
  xfrm <- self$find(qn("p:xfrm")); if (is.null(xfrm)) Emu(0L) else xfrm$y %||% Emu(0L)
}, overwrite = TRUE)

CT_GraphicalObjectFrame$set("active", "cx", function(value) {
  if (!missing(value)) {
    .gf_get_or_add_pxfrm(self)$cx <- value; return(invisible(value))
  }
  xfrm <- self$find(qn("p:xfrm")); if (is.null(xfrm)) Emu(0L) else xfrm$cx %||% Emu(0L)
}, overwrite = TRUE)

CT_GraphicalObjectFrame$set("active", "cy", function(value) {
  if (!missing(value)) {
    .gf_get_or_add_pxfrm(self)$cy <- value; return(invisible(value))
  }
  xfrm <- self$find(qn("p:xfrm")); if (is.null(xfrm)) Emu(0L) else xfrm$cy %||% Emu(0L)
}, overwrite = TRUE)

# Add tbl accessor to CT_GraphicalObjectFrame
CT_GraphicalObjectFrame$set("active", "tbl", function() {
  r <- xml2::xml_find_first(
    private$.node,
    "a:graphic/a:graphicData/a:tbl",
    ns = c(a = .nsmap[["a"]])
  )
  if (inherits(r, "xml_missing")) return(NULL)
  wrap_element(r)
})


# ============================================================================
# Factory — create a graphicFrame containing a table
# ============================================================================

#' Create a new <p:graphicFrame> wrapping a table
#' @noRd
CT_GraphicalObjectFrame_new_table_graphicFrame <- function(id, name, rows, cols,
                                                            x, y, cx, cy) {
  p <- .nsmap[["p"]]; a <- .nsmap[["a"]]
  xml_str <- sprintf(paste0(
    '<p:graphicFrame xmlns:p="%s" xmlns:a="%s">',
      '<p:nvGraphicFramePr>',
        '<p:cNvPr id="%d" name="%s"/>',
        '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>',
        '<p:nvPr/>',
      '</p:nvGraphicFramePr>',
      '<p:xfrm>',
        '<a:off x="%d" y="%d"/>',
        '<a:ext cx="%d" cy="%d"/>',
      '</p:xfrm>',
      '<a:graphic>',
        '<a:graphicData uri="%s"/>',
      '</a:graphic>',
    '</p:graphicFrame>'
  ), p, a,
    as.integer(id), as.character(name),
    as.integer(x), as.integer(y), as.integer(cx), as.integer(cy),
    .GRAPHIC_DATA_URI_TABLE)

  gf <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))

  # Build the table and append it to graphicData
  tbl <- CT_Table_new_tbl(rows, cols, cx, cy)
  graphicData_node <- xml2::xml_find_first(
    gf$get_node(), "a:graphic/a:graphicData",
    ns = c(a = a)
  )
  xml2::xml_add_child(graphicData_node, tbl$get_node())

  gf
}


# ============================================================================
# Element registration
# ============================================================================

.onLoad_oxml_table <- function() {
  register_element_cls("a:tblPr",    CT_TableProperties)
  register_element_cls("a:tblGrid",  CT_TableGrid)
  register_element_cls("a:gridCol",  CT_TableCol)
  register_element_cls("a:tcPr",     CT_TableCellProperties)
  register_element_cls("a:tc",       CT_TableCell)
  register_element_cls("a:tr",       CT_TableRow)
  register_element_cls("a:tbl",      CT_Table)
  # Also register p:xfrm for graphicFrame transforms (same structure as a:xfrm)
  register_element_cls("p:xfrm",     CT_Transform2D)
}
