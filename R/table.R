# Table domain objects.
#
# Ported from python-pptx/src/pptx/table.py.
# Provides Table, TableRow, TableColumn, TableCell and their collections.

# ============================================================================
# Helper — find (row, col) position of a CT_TableCell in a Table
# ============================================================================

.find_cell_pos <- function(table, tc_elm) {
  tbl <- table$.__enclos_env__$private$.tbl
  rows <- tbl$tr_lst
  for (r in seq_along(rows)) {
    tcs <- rows[[r]]$tc_lst
    for (c in seq_along(tcs)) {
      if (identical(tcs[[c]]$get_node(), tc_elm$get_node())) {
        return(c(r, c))
      }
    }
  }
  stop("cell not found in table", call. = FALSE)
}


# ============================================================================
# TableCell — wraps a <a:tc> element
# ============================================================================

#' Table cell proxy
#'
#' Provides access to cell text, fill, margins, and merge state.
#' Access via `table$cell(row, col)` or `table$rows[[i]]$cells[[j]]`.
#'
#' @keywords internal
#' @export
TableCell <- R6::R6Class(
  "TableCell",

  public = list(
    initialize = function(tc, parent) {
      private$.tc     <- tc
      private$.parent <- parent  # Table
    },

    # Merge a rectangular range from this cell to other_cell (bottom-right corner).
    # After merging, only this cell (the origin) retains content; spanned cells
    # get hMerge / vMerge flags set.
    merge = function(other_cell) {
      tbl <- private$.parent
      # Find (row, col) of both cells by iterating the table
      self_pos  <- .find_cell_pos(tbl, private$.tc)
      other_pos <- .find_cell_pos(tbl, other_cell$.__enclos_env__$private$.tc)
      r1 <- min(self_pos[1], other_pos[1]); r2 <- max(self_pos[1], other_pos[1])
      c1 <- min(self_pos[2], other_pos[2]); c2 <- max(self_pos[2], other_pos[2])
      n_rows <- r2 - r1 + 1L; n_cols <- c2 - c1 + 1L
      # Configure all cells in the range
      for (r in r1:r2) {
        for (c in c1:c2) {
          tc <- tbl$.__enclos_env__$private$.tbl$tc(r, c)
          if (r == r1 && c == c1) {
            # Origin cell
            if (n_cols > 1L) tc$gridSpan <- n_cols else tc$gridSpan <- 1L
            if (n_rows > 1L) tc$rowSpan  <- n_rows else tc$rowSpan  <- 1L
            tc$hMerge <- FALSE; tc$vMerge <- FALSE
          } else if (r == r1) {
            # Same row, non-origin
            tc$hMerge <- TRUE; tc$vMerge <- FALSE
            tc$gridSpan <- 1L; tc$rowSpan <- 1L
          } else if (c == c1) {
            # Same col, non-origin
            tc$vMerge <- TRUE; tc$hMerge <- FALSE
            tc$gridSpan <- 1L; tc$rowSpan <- 1L
          } else {
            # Interior
            tc$hMerge <- TRUE; tc$vMerge <- TRUE
            tc$gridSpan <- 1L; tc$rowSpan <- 1L
          }
        }
      }
      invisible(self)
    }
  ),

  active = list(
    # The parent Table
    part = function() private$.parent$part,

    # Cell text (read/write) — convenience shortcut for text_frame$text
    text = function(value) {
      if (!missing(value)) {
        txBody <- private$.tc$get_or_add_txBody()
        tf <- TextFrame$new(txBody, self)
        tf$text <- value
        return(invisible(value))
      }
      txBody <- private$.tc$txBody
      if (is.null(txBody)) return("")
      TextFrame$new(txBody, self)$text
    },

    # TextFrame for full paragraph/run control.
    # No-op setter handles write-back from chaining.
    text_frame = function(value) {
      if (!missing(value)) return(invisible(NULL))
      txBody <- private$.tc$get_or_add_txBody()
      TextFrame$new(txBody, self)
    },

    # FillFormat for cell background.
    # No-op setter handles write-back from chaining.
    fill = function(value) {
      if (!missing(value)) return(invisible(NULL))
      tcPr <- private$.tc$get_or_add_tcPr()
      FillFormat$new(tcPr)
    },

    # Left margin in EMU (read/write). NULL uses PowerPoint default.
    margin_left = function(value) {
      if (!missing(value)) {
        private$.tc$get_or_add_tcPr()$marL <- value
        return(invisible(value))
      }
      tcPr <- private$.tc$tcPr
      if (is.null(tcPr)) return(NULL)
      tcPr$marL
    },

    # Right margin in EMU (read/write).
    margin_right = function(value) {
      if (!missing(value)) {
        private$.tc$get_or_add_tcPr()$marR <- value
        return(invisible(value))
      }
      tcPr <- private$.tc$tcPr
      if (is.null(tcPr)) return(NULL)
      tcPr$marR
    },

    # Top margin in EMU (read/write).
    margin_top = function(value) {
      if (!missing(value)) {
        private$.tc$get_or_add_tcPr()$marT <- value
        return(invisible(value))
      }
      tcPr <- private$.tc$tcPr
      if (is.null(tcPr)) return(NULL)
      tcPr$marT
    },

    # Bottom margin in EMU (read/write).
    margin_bottom = function(value) {
      if (!missing(value)) {
        private$.tc$get_or_add_tcPr()$marB <- value
        return(invisible(value))
      }
      tcPr <- private$.tc$tcPr
      if (is.null(tcPr)) return(NULL)
      tcPr$marB
    },

    # TRUE if this is the top-left origin of a merged range
    is_merge_origin = function() private$.tc$is_merge_origin,

    # TRUE if this cell is spanned (part of merge, not the origin)
    is_spanned = function() private$.tc$is_spanned,

    # Number of columns this merge origin spans (only valid on merge-origin cells)
    span_width = function() private$.tc$gridSpan,

    # Number of rows this merge origin spans (only valid on merge-origin cells)
    span_height = function() private$.tc$rowSpan
  ),

  private = list(.tc = NULL, .parent = NULL)
)


# ============================================================================
# TableCells — cell collection for a single row
# ============================================================================

#' Cell collection for a table row
#' @keywords internal
#' @export
TableCells <- R6::R6Class(
  "TableCells",

  public = list(
    initialize = function(tr, parent) {
      private$.tr     <- tr
      private$.parent <- parent  # Table
    },

    get = function(idx) {
      tc_list <- private$.tr$tc_lst
      if (idx < 1L || idx > length(tc_list)) {
        stop("cell index out of range", call. = FALSE)
      }
      TableCell$new(tc_list[[idx]], private$.parent)
    },

    to_list = function() lapply(seq_len(length(self)), self$get)
  ),

  private = list(.tr = NULL, .parent = NULL)
)

#' @export
length.TableCells <- function(x) {
  length(x$.__enclos_env__$private$.tr$tc_lst)
}

#' @export
`[[.TableCells` <- function(x, i) x$get(i)

#' @export
`[[<-.TableCells` <- function(x, i, value) x


# ============================================================================
# TableRow — wraps a <a:tr> element
# ============================================================================

#' Table row proxy
#'
#' Provides access to row height and cells.
#' Access via `table$rows[[i]]`.
#'
#' @keywords internal
#' @export
TableRow <- R6::R6Class(
  "TableRow",

  public = list(
    initialize = function(tr, parent) {
      private$.tr     <- tr
      private$.parent <- parent  # Table
    }
  ),

  active = list(
    # Row height in EMU (read/write)
    height = function(value) {
      if (!missing(value)) { private$.tr$h <- value; return(invisible(value)) }
      private$.tr$h
    },

    # TableCells collection for cells in this row
    cells = function(value) {
      if (!missing(value)) return(invisible(NULL))
      TableCells$new(private$.tr, private$.parent)
    }
  ),

  private = list(.tr = NULL, .parent = NULL)
)


# ============================================================================
# TableRows — collection of TableRow objects
# ============================================================================

#' Row collection for a table
#' @keywords internal
#' @export
TableRows <- R6::R6Class(
  "TableRows",

  public = list(
    initialize = function(tbl, parent) {
      private$.tbl    <- tbl
      private$.parent <- parent  # Table
    },

    get = function(idx) {
      tr_list <- private$.tbl$tr_lst
      if (idx < 1L || idx > length(tr_list)) {
        stop("row index out of range", call. = FALSE)
      }
      TableRow$new(tr_list[[idx]], private$.parent)
    },

    to_list = function() lapply(seq_len(length(self)), self$get)
  ),

  private = list(.tbl = NULL, .parent = NULL)
)

#' @export
length.TableRows <- function(x) {
  length(x$.__enclos_env__$private$.tbl$tr_lst)
}

#' @export
`[[.TableRows` <- function(x, i) x$get(i)

#' @export
`[[<-.TableRows` <- function(x, i, value) x


# ============================================================================
# TableColumn — wraps a <a:gridCol> element
# ============================================================================

#' Table column proxy
#'
#' Provides access to column width.
#' Access via `table$columns[[i]]`.
#'
#' @keywords internal
#' @export
TableColumn <- R6::R6Class(
  "TableColumn",

  public = list(
    initialize = function(gridCol, parent) {
      private$.gridCol <- gridCol
      private$.parent  <- parent  # Table
    }
  ),

  active = list(
    # Column width in EMU (read/write)
    width = function(value) {
      if (!missing(value)) { private$.gridCol$w <- value; return(invisible(value)) }
      private$.gridCol$w
    }
  ),

  private = list(.gridCol = NULL, .parent = NULL)
)


# ============================================================================
# TableColumns — collection of TableColumn objects
# ============================================================================

#' Column collection for a table
#' @keywords internal
#' @export
TableColumns <- R6::R6Class(
  "TableColumns",

  public = list(
    initialize = function(tbl, parent) {
      private$.tbl    <- tbl
      private$.parent <- parent  # Table
    },

    get = function(idx) {
      gc_list <- private$.tbl$gridCol_lst
      if (idx < 1L || idx > length(gc_list)) {
        stop("column index out of range", call. = FALSE)
      }
      TableColumn$new(gc_list[[idx]], private$.parent)
    },

    to_list = function() lapply(seq_len(length(self)), self$get)
  ),

  private = list(.tbl = NULL, .parent = NULL)
)

#' @export
length.TableColumns <- function(x) {
  length(x$.__enclos_env__$private$.tbl$gridCol_lst)
}

#' @export
`[[.TableColumns` <- function(x, i) x$get(i)

#' @export
`[[<-.TableColumns` <- function(x, i, value) x


# ============================================================================
# Table — domain object wrapping <a:tbl>
# ============================================================================

#' Table proxy
#'
#' Provides access to table cells, rows, columns, and formatting flags.
#' Access via `graphic_frame$table`.
#'
#' @keywords internal
#' @export
Table <- R6::R6Class(
  "Table",

  public = list(
    initialize = function(tbl_elm, graphic_frame) {
      private$.tbl <- tbl_elm
      private$.gf  <- graphic_frame
    },

    # Return the TableCell at 1-based (row_idx, col_idx)
    cell = function(row_idx, col_idx) {
      tc <- private$.tbl$tc(row_idx, col_idx)
      TableCell$new(tc, self)
    },

    # Return all cells as a list (left-to-right, top-to-bottom)
    iter_cells = function() {
      lapply(private$.tbl$iter_tcs(), function(tc) TableCell$new(tc, self))
    }
  ),

  active = list(
    # TableRows collection
    rows = function(value) {
      if (!missing(value)) return(invisible(NULL))
      TableRows$new(private$.tbl, self)
    },

    # TableColumns collection
    columns = function(value) {
      if (!missing(value)) return(invisible(NULL))
      TableColumns$new(private$.tbl, self)
    },

    # Slide part (for TextFrame compatibility)
    part = function() private$.gf$part,

    # Boolean table-style flags (read/write)
    first_row = function(value) {
      tblPr <- private$.tbl$tblPr
      if (is.null(tblPr)) return(FALSE)
      if (!missing(value)) { tblPr$firstRow <- value; return(invisible(value)) }
      isTRUE(tblPr$firstRow)
    },

    last_row = function(value) {
      tblPr <- private$.tbl$tblPr
      if (is.null(tblPr)) return(FALSE)
      if (!missing(value)) { tblPr$lastRow <- value; return(invisible(value)) }
      isTRUE(tblPr$lastRow)
    },

    first_col = function(value) {
      tblPr <- private$.tbl$tblPr
      if (is.null(tblPr)) return(FALSE)
      if (!missing(value)) { tblPr$firstCol <- value; return(invisible(value)) }
      isTRUE(tblPr$firstCol)
    },

    last_col = function(value) {
      tblPr <- private$.tbl$tblPr
      if (is.null(tblPr)) return(FALSE)
      if (!missing(value)) { tblPr$lastCol <- value; return(invisible(value)) }
      isTRUE(tblPr$lastCol)
    },

    horz_banding = function(value) {
      tblPr <- private$.tbl$tblPr
      if (is.null(tblPr)) return(FALSE)
      if (!missing(value)) { tblPr$bandRow <- value; return(invisible(value)) }
      isTRUE(tblPr$bandRow)
    },

    vert_banding = function(value) {
      tblPr <- private$.tbl$tblPr
      if (is.null(tblPr)) return(FALSE)
      if (!missing(value)) { tblPr$bandCol <- value; return(invisible(value)) }
      isTRUE(tblPr$bandCol)
    },

    # Table style GUID string, e.g. "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}".
    # NULL when no tblPr or tableStyleId element exists.
    style_id = function(value) {
      tblPr <- private$.tbl$tblPr
      if (is.null(tblPr)) {
        if (!missing(value)) {
          stop("cannot set style_id: table has no tblPr element", call. = FALSE)
        }
        return(NULL)
      }
      if (!missing(value)) {
        ts_elm <- tblPr$get_or_add_tableStyleId()
        xml2::xml_set_text(ts_elm$get_node(), as.character(value))
        return(invisible(value))
      }
      ts_elm <- tblPr$tableStyleId
      if (is.null(ts_elm)) return(NULL)
      ts_elm$text
    }
  ),

  private = list(.tbl = NULL, .gf = NULL)
)
