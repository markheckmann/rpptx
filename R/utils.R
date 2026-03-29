# Utility functions and classes.
#
# Length classes for EMU (English Metric Unit) conversions, mirroring
# python-pptx's util.py. All lengths are stored internally as integer EMU
# values and can be converted to other units via helper functions.

# --- Constants ---
.EMUS_PER_INCH <- 914400L
.EMUS_PER_CENTIPOINT <- 127L
.EMUS_PER_CM <- 360000L
.EMUS_PER_MM <- 36000L
.EMUS_PER_PT <- 12700L


#' Create a Length value in English Metric Units (EMU)
#'
#' Base constructor for length values. All length values are stored as integer
#' EMU values. Use convenience constructors [Inches()], [Cm()], [Mm()],
#' [Pt()], [Emu()], and [Centipoints()] for common units.
#'
#' @param emu Integer EMU value.
#' @return An integer with class `"Length"`.
#' @export
Length <- function(emu) {
  emu <- as.integer(emu)
  structure(emu, class = c("Length", "integer"))
}


#' @rdname Length
#' @param inches Numeric length in inches.
#' @export
Inches <- function(inches) {
  Length(as.integer(round(inches * .EMUS_PER_INCH)))
}


#' @rdname Length
#' @param cm Numeric length in centimeters.
#' @export
Cm <- function(cm) {
  Length(as.integer(round(cm * .EMUS_PER_CM)))
}


#' @rdname Length
#' @param mm Numeric length in millimeters.
#' @export
Mm <- function(mm) {
  Length(as.integer(round(mm * .EMUS_PER_MM)))
}


#' @rdname Length
#' @param pt Numeric length in points.
#' @export
Pt <- function(pt) {
  Length(as.integer(round(pt * .EMUS_PER_PT)))
}


#' @rdname Length
#' @export
Emu <- function(emu) {
  Length(as.integer(emu))
}


#' @rdname Length
#' @param centipoints Integer length in hundredths of a point (1/7200 inch).
#' @export
Centipoints <- function(centipoints) {
  Length(as.integer(round(centipoints * .EMUS_PER_CENTIPOINT)))
}


# --- Conversion helpers ---

#' Convert a Length value to inches
#' @param x A Length value (integer EMU).
#' @return Numeric value in the target unit.
#' @export
as_inches <- function(x) as.integer(x) / .EMUS_PER_INCH

#' @rdname as_inches
#' @export
as_cm <- function(x) as.integer(x) / .EMUS_PER_CM

#' @rdname as_inches
#' @export
as_mm <- function(x) as.integer(x) / .EMUS_PER_MM

#' @rdname as_inches
#' @export
as_pt <- function(x) as.integer(x) / .EMUS_PER_PT

#' @rdname as_inches
#' @export
as_emu <- function(x) as.integer(x)

#' @rdname as_inches
#' @export
as_centipoints <- function(x) as.integer(x) %/% .EMUS_PER_CENTIPOINT


# --- S3 methods ---

#' @export
print.Length <- function(x, ...) {
  cat(sprintf("<Length: %d EMU (%.2f in, %.2f cm, %.1f pt)>\n",
              as.integer(x), as_inches(x), as_cm(x), as_pt(x)))
  invisible(x)
}

#' @export
format.Length <- function(x, ...) {
  sprintf("%d", as.integer(x))
}


# --- Lazy active binding helper for R6 ---

#' Create a lazy active binding function for R6 classes
#'
#' Returns a function suitable for use in R6 `active` list that evaluates
#' `fn` on first access and caches the result in private storage.
#'
#' @param fn A function taking `self` that computes the value.
#' @param cache_field Name of the private field to use for caching.
#' @return A function suitable for R6 active bindings.
#' @noRd
lazy_active_binding <- function(fn, cache_field) {
  # R6 replaces active binding function environments, so closures break.
  # We bake the values directly into the function AST using bquote().
  f <- eval(bquote(function(value) {
    if (!missing(value)) {
      stop("This property is read-only.", call. = FALSE)
    }
    cached <- private[[.(cache_field)]]
    if (is.null(cached)) {
      cached <- (.(fn))(self)
      private[[.(cache_field)]] <- cached
    }
    cached
  }))
  f
}
