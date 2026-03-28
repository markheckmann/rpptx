# Presentation — user-facing presentation object.
#
# Ported from python-pptx/src/pptx/presentation.py. Provides the
# Presentation R6 class returned by the Presentation() API function.

#' Presentation object
#'
#' User-facing object for a PowerPoint presentation. Wraps the
#' CT_Presentation XML element and delegates to PresentationPart.
#'
#' @include shared.R
#' @keywords internal
#' @export
Presentation <- R6::R6Class(
  "Presentation",
  inherit = PartElementProxy,

  public = list(
    # Save presentation to a file
    save = function(file) {
      self$part$save(file)
    }
  ),

  active = list(
    # Slide width in EMU (read/write)
    slide_width = function(value) {
      if (!missing(value)) {
        sld_sz <- self$element$get_or_add_sldSz()
        sld_sz$cx <- value
        return(invisible(value))
      }
      sld_sz <- self$element$sldSz
      if (is.null(sld_sz)) return(NULL)
      sld_sz$cx
    },

    # Slide height in EMU (read/write)
    slide_height = function(value) {
      if (!missing(value)) {
        sld_sz <- self$element$get_or_add_sldSz()
        sld_sz$cy <- value
        return(invisible(value))
      }
      sld_sz <- self$element$sldSz
      if (is.null(sld_sz)) return(NULL)
      sld_sz$cy
    }
  )
)
