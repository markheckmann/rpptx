# DML-related enumeration constants.
#
# Ported from python-pptx/src/pptx/enum/dml.py.


#' MSO_THEME_COLOR — theme color XML value strings
#'
#' String values for the `val` attribute on `<a:schemeClr>`. Assign to
#' `color_format$theme_color`.
#'
#' @noRd
#' @export
MSO_THEME_COLOR <- list(
  NOT_THEME_COLOR    = NULL,
  DARK_1             = "dk1",
  LIGHT_1            = "lt1",
  DARK_2             = "dk2",
  LIGHT_2            = "lt2",
  ACCENT_1           = "accent1",
  ACCENT_2           = "accent2",
  ACCENT_3           = "accent3",
  ACCENT_4           = "accent4",
  ACCENT_5           = "accent5",
  ACCENT_6           = "accent6",
  HYPERLINK          = "hlink",
  FOLLOWED_HYPERLINK = "folHlink"
)


#' MSO_COLOR_TYPE — color source type strings
#'
#' @noRd
#' @export
MSO_COLOR_TYPE <- list(
  RGB    = "rgb",
  SCHEME = "scheme",
  SCRGB  = "scrgb",
  HSL    = "hsl",
  PRESET = "preset",
  SYSTEM = "system"
)


#' MSO_FILL — fill type strings
#'
#' @noRd
#' @export
MSO_FILL <- list(
  BACKGROUND = "background",
  GRADIENT   = "gradient",
  GROUP      = "group",
  PATTERNED  = "patterned",
  PICTURE    = "picture",
  SOLID      = "solid"
)


#' MSO_LINE_DASH_STYLE — preset dash pattern strings for `<a:prstDash val="..."/>`
#'
#' @noRd
#' @export
MSO_LINE_DASH_STYLE <- list(
  DASH          = "dash",
  DASH_DOT      = "dashDot",
  DASH_DOT_DOT  = "lgDashDotDot",
  LONG_DASH     = "lgDash",
  LONG_DASH_DOT = "lgDashDot",
  ROUND_DOT     = "sysDot",
  SOLID         = "solid",
  SQUARE_DOT    = "sysDash"
)


#' MSO_PATTERN_TYPE — preset pattern fill type strings
#'
#' @noRd
#' @export
MSO_PATTERN_TYPE <- list(
  CROSS                    = "cross",
  DARK_DOWNWARD_DIAGONAL   = "dkDnDiag",
  DARK_HORIZONTAL          = "dkHorz",
  DARK_UPWARD_DIAGONAL     = "dkUpDiag",
  DARK_VERTICAL            = "dkVert",
  DASHED_DOWNWARD_DIAGONAL = "dashDnDiag",
  DASHED_HORIZONTAL        = "dashHorz",
  DASHED_UPWARD_DIAGONAL   = "dashUpDiag",
  DASHED_VERTICAL          = "dashVert",
  DIAGONAL_BRICK           = "diagBrick",
  DIAGONAL_CROSS           = "diagCross",
  DIVOT                    = "divot",
  DOTTED_DIAMOND           = "dotDmnd",
  DOTTED_GRID              = "dotGrid",
  DOWNWARD_DIAGONAL        = "dnDiag",
  HORIZONTAL               = "horz",
  HORIZONTAL_BRICK         = "horzBrick",
  LARGE_CHECKER_BOARD      = "lgCheck",
  LARGE_CONFETTI           = "lgConfetti",
  LARGE_GRID               = "lgGrid",
  LIGHT_DOWNWARD_DIAGONAL  = "ltDnDiag",
  LIGHT_HORIZONTAL         = "ltHorz",
  LIGHT_UPWARD_DIAGONAL    = "ltUpDiag",
  LIGHT_VERTICAL           = "ltVert",
  NARROW_HORIZONTAL        = "narHorz",
  NARROW_VERTICAL          = "narVert",
  OUTLINED_DIAMOND         = "openDmnd",
  PLAID                    = "plaid",
  SHINGLE                  = "shingle",
  SMALL_CHECKER_BOARD      = "smCheck",
  SMALL_CONFETTI           = "smConfetti",
  SMALL_GRID               = "smGrid",
  SOLID_DIAMOND            = "solidDmnd",
  SPHERE                   = "sphere",
  TRELLIS                  = "trellis",
  UPWARD_DIAGONAL          = "upDiag",
  VERTICAL                 = "vert",
  WAVE                     = "wave",
  WEAVE                    = "weave",
  WIDE_DOWNWARD_DIAGONAL   = "wdDnDiag",
  WIDE_UPWARD_DIAGONAL     = "wdUpDiag",
  ZIG_ZAG                  = "zigZag"
)
