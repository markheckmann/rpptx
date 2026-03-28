# Text-related enumeration constants.
#
# Ported from python-pptx/src/pptx/enum/text.py.


#' PP_PARAGRAPH_ALIGNMENT — paragraph alignment
#'
#' XML string values for `a:pPr/@algn`. Use with paragraph `alignment` property.
#'
#' @keywords internal
#' @export
PP_PARAGRAPH_ALIGNMENT <- list(
  CENTER           = "ctr",
  DISTRIBUTE       = "dist",
  JUSTIFY          = "just",
  JUSTIFY_LOW      = "justLow",
  LEFT             = "l",
  RIGHT            = "r",
  THAI_DISTRIBUTE  = "thaiDist",
  MIXED            = ""
)

#' @keywords internal
#' @export
PP_ALIGN <- PP_PARAGRAPH_ALIGNMENT


#' MSO_AUTO_SIZE — text-box auto-sizing behaviour
#'
#' Integer constants for `a:bodyPr/@autofit` / spAutoFit / normAutofit.
#'
#' @keywords internal
#' @export
MSO_AUTO_SIZE <- list(
  NONE             = 0L,
  SHAPE_TO_FIT_TEXT = 1L,
  TEXT_TO_FIT_SHAPE = 2L,
  MIXED            = -2L
)


#' MSO_VERTICAL_ANCHOR — vertical text anchor
#'
#' XML string values for `a:bodyPr/@anchor`. Use with `text_frame$vertical_anchor`.
#'
#' @keywords internal
#' @export
MSO_VERTICAL_ANCHOR <- list(
  TOP   = "t",
  MIDDLE = "ctr",
  BOTTOM = "b",
  MIXED  = ""
)
