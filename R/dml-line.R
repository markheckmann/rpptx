# DML line (border) domain objects.
#
# Ported from python-pptx/src/pptx/dml/line.py.
# Provides LineFormat R6 class.

# ============================================================================
# LineFormat — access to line/border properties on a shape
# ============================================================================

#' Line format accessor
#'
#' Wraps the shape-properties element (`<p:spPr>`) and exposes border/line
#' properties. Access via `shape$line`.
#'
#' @keywords internal
#' @export
LineFormat <- R6::R6Class(
  "LineFormat",

  public = list(
    #' @param spPr A CT_ShapeProperties element.
    initialize = function(spPr) {
      private$.spPr <- spPr
    }
  ),

  active = list(
    # CT_LineProperties element, creating it if absent.
    # No-op setter handles R's write-back when chaining: shape$line$width <- val
    .ln = function(value) {
      if (!missing(value)) return(invisible(NULL))
      spPr <- private$.spPr
      ln <- spPr$ln
      if (!is.null(ln)) return(ln)
      spPr$get_or_add_ln()
    },

    # Line width in EMU (read/write). NULL means no explicit width (theme default).
    width = function(value) {
      if (!missing(value)) {
        self$.ln$w <- value
        return(invisible(value))
      }
      ln <- private$.spPr$ln
      if (is.null(ln)) return(NULL)
      ln$w
    },

    # Dash style string (MSO_LINE_DASH_STYLE value). NULL if not set.
    dash_style = function(value) {
      if (!missing(value)) {
        prstDash <- self$.ln$get_or_add_prstDash()
        prstDash$val <- value
        return(invisible(value))
      }
      ln <- private$.spPr$ln
      if (is.null(ln)) return(NULL)
      pd <- ln$prstDash
      if (is.null(pd)) return(NULL)
      pd$val
    },

    # ColorFormat for the line color.
    # No-op setter allows chaining: shape$line$color$rgb <- RGBColor(...)
    color = function(value) {
      if (!missing(value)) return(invisible(NULL))
      ln <- self$.ln
      fill_elm <- ln$eg_fillProperties
      if (is.null(fill_elm) || !inherits(fill_elm, "CT_SolidColorFillProperties")) {
        # Auto-apply solid fill so a color can be set
        fill_elm <- ln$get_or_add_solidFill()
      }
      ColorFormat$new(fill_elm)
    },

    # TRUE if line has no fill (transparent/no border)
    fill = function(value) {
      if (!missing(value)) return(invisible(NULL))
      ln <- private$.spPr$ln
      if (is.null(ln)) return(NULL)
      fill_elm <- ln$eg_fillProperties
      if (is.null(fill_elm)) return(NULL)
      if (inherits(fill_elm, "CT_NoFill")) return(MSO_FILL$BACKGROUND)
      if (inherits(fill_elm, "CT_SolidColorFillProperties")) return(MSO_FILL$SOLID)
      NULL
    }
  ),

  private = list(.spPr = NULL)
)
