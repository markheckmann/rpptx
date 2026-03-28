# Print / format methods for rpptx objects.
#
# Defines S3 format() and print() generics for Presentation, Slide,
# SlideLayout, SlideMaster, shape proxies, Table, and TextFrame.
# Also provides as.data.frame.Table.


# ============================================================================
# Internal helpers
# ============================================================================

# Convert EMU to inches as a rounded string.
.emu_in <- function(emu) round(as.numeric(emu) / 914400, 2)

# Map MSO_SHAPE_TYPE integer to a short label.
.shape_type_label <- function(type) {
  switch(as.character(type),
    "1"  = "AutoShape",
    "3"  = "Chart",
    "6"  = "Group",
    "9"  = "Line",
    "13" = "Picture",
    "14" = "Placeholder",
    "17" = "TextBox",
    "19" = "Table",
    "25" = "Media",
    sprintf("type=%s", type)
  )
}

# Truncate a string for display.
.trunc <- function(s, n = 30) {
  if (is.null(s) || is.na(s)) return("")
  s <- gsub("[\n\r]+", " ", s)
  if (nchar(s) > n) paste0(substr(s, 1, n - 1), "\u2026") else s
}


# ============================================================================
# Presentation
# ============================================================================

#' @export
format.Presentation <- function(x, ...) {
  ns <- length(x$slides)
  nl <- length(x$slide_layouts)
  nm <- length(x$slide_masters)
  w  <- .emu_in(x$slide_width)
  h  <- .emu_in(x$slide_height)
  paste0(
    "<Presentation>\n",
    sprintf("  slides  : %d\n", ns),
    sprintf("  size    : %.2f \u00d7 %.2f in\n", w, h),
    sprintf("  layouts : %d\n", nl),
    sprintf("  masters : %d", nm)
  )
}

#' @export
print.Presentation <- function(x, ...) {
  cat(format(x, ...), "\n")
  invisible(x)
}


# ============================================================================
# Slide
# ============================================================================

#' @export
format.Slide <- function(x, ...) {
  shapes <- tryCatch(x$shapes$to_list(), error = function(e) list())
  n <- length(shapes)
  header <- sprintf("<Slide>  %d shape%s", n, if (n == 1L) "" else "s")
  if (n == 0L) return(header)
  rows <- vapply(shapes, function(s) {
    lbl  <- .shape_type_label(tryCatch(s$shape_type, error = function(e) "?"))
    nm   <- tryCatch(.trunc(s$name, 24), error = function(e) "")
    l    <- tryCatch(.emu_in(s$left),   error = function(e) NA)
    t    <- tryCatch(.emu_in(s$top),    error = function(e) NA)
    w    <- tryCatch(.emu_in(s$width),  error = function(e) NA)
    h    <- tryCatch(.emu_in(s$height), error = function(e) NA)
    sprintf("  %-14s  %-24s  %.2f\u00d7%.2f in @ %.2f, %.2f", lbl, nm, w, h, l, t)
  }, character(1))
  paste(c(header, rows), collapse = "\n")
}

#' @export
print.Slide <- function(x, ...) {
  cat(format(x, ...), "\n")
  invisible(x)
}


# ============================================================================
# SlideLayout
# ============================================================================

#' @export
format.SlideLayout <- function(x, ...) {
  nm <- tryCatch(x$name, error = function(e) "")
  phs <- tryCatch(length(x$placeholders$to_list()), error = function(e) 0L)
  sprintf("<SlideLayout>  \"%s\"  (%d placeholder%s)", nm, phs,
          if (phs == 1L) "" else "s")
}

#' @export
print.SlideLayout <- function(x, ...) {
  cat(format(x, ...), "\n")
  invisible(x)
}


# ============================================================================
# SlideMaster
# ============================================================================

#' @export
format.SlideMaster <- function(x, ...) {
  nl  <- tryCatch(length(x$slide_layouts), error = function(e) 0L)
  phs <- tryCatch(length(x$placeholders$to_list()), error = function(e) 0L)
  sprintf("<SlideMaster>  %d layout%s  |  %d placeholder%s",
          nl,  if (nl  == 1L) "" else "s",
          phs, if (phs == 1L) "" else "s")
}

#' @export
print.SlideMaster <- function(x, ...) {
  cat(format(x, ...), "\n")
  invisible(x)
}


# ============================================================================
# Shapes — BaseShape covers all shape subclasses
# ============================================================================

#' @export
format.BaseShape <- function(x, ...) {
  lbl <- tryCatch(.shape_type_label(x$shape_type), error = function(e) "Shape")
  nm  <- tryCatch(.trunc(x$name, 30), error = function(e) "")
  l   <- tryCatch(.emu_in(x$left),   error = function(e) NA)
  t   <- tryCatch(.emu_in(x$top),    error = function(e) NA)
  w   <- tryCatch(.emu_in(x$width),  error = function(e) NA)
  h   <- tryCatch(.emu_in(x$height), error = function(e) NA)
  sprintf("<%s>  \"%s\"  %.2f\u00d7%.2f in  @ %.2f, %.2f in", lbl, nm, w, h, l, t)
}

#' @export
print.BaseShape <- function(x, ...) {
  cat(format(x, ...), "\n")
  invisible(x)
}

# Subclasses inherit format/print from BaseShape via S3 dispatch through
# the R6 class vector: SlidePlaceholder → Shape → BaseShape → R6


# ============================================================================
# Table
# ============================================================================

#' @export
format.Table <- function(x, ...) {
  nr <- tryCatch(length(x$rows),    error = function(e) "?")
  nc <- tryCatch(length(x$columns), error = function(e) "?")
  sprintf("<Table>  %s rows \u00d7 %s cols", nr, nc)
}

#' @export
print.Table <- function(x, ...) {
  cat(format(x, ...), "\n")
  invisible(x)
}

#' Convert a Table to a data frame of cell text values
#'
#' Each cell's text content becomes one element of the data frame.
#' Row 1 becomes the first row of the data frame (not used as column names).
#' To promote the first row to headers use `setNames(as.data.frame(tbl), as.data.frame(tbl)[1, ])`.
#'
#' @param x A `Table` object.
#' @param row.names NULL or a character vector of row names (default: NULL).
#' @param optional Ignored; included for S3 compatibility.
#' @param ... Ignored.
#' @return A data.frame with character columns named V1, V2, …
#' @export
as.data.frame.Table <- function(x, row.names = NULL, optional = FALSE, ...) {
  nr <- length(x$rows)
  nc <- length(x$columns)
  mat <- matrix(NA_character_, nrow = nr, ncol = nc)
  for (i in seq_len(nr)) {
    for (j in seq_len(nc)) {
      mat[i, j] <- tryCatch(x$cell(i, j)$text, error = function(e) NA_character_)
    }
  }
  df <- as.data.frame(mat, stringsAsFactors = FALSE)
  colnames(df) <- paste0("V", seq_len(nc))
  if (!is.null(row.names)) rownames(df) <- row.names
  df
}


# ============================================================================
# TextFrame
# ============================================================================

#' @export
format.TextFrame <- function(x, ...) {
  txt <- tryCatch(.trunc(x$text, 40), error = function(e) "")
  sprintf("<TextFrame>  \"%s\"", txt)
}

#' @export
print.TextFrame <- function(x, ...) {
  cat(format(x, ...), "\n")
  invisible(x)
}


# ============================================================================
# PlaceholderFormat
# ============================================================================

#' @export
format.PlaceholderFormat <- function(x, ...) {
  sprintf("<PlaceholderFormat>  idx=%s  type=%s",
          tryCatch(x$idx, error = function(e) "?"),
          tryCatch(x$type, error = function(e) "?"))
}

#' @export
print.PlaceholderFormat <- function(x, ...) {
  cat(format(x, ...), "\n")
  invisible(x)
}
