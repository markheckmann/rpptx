# Shape-related enumeration constants.
#
# Ported from python-pptx/src/pptx/enum/shapes.py.


#' MSO_SHAPE_TYPE — shape type classification constants
#'
#' Integer constants classifying shapes by type, mirroring MSO_SHAPE_TYPE.
#'
#' @noRd
#' @export
MSO_SHAPE_TYPE <- list(
  AUTO_SHAPE      = 1L,
  CALLOUT         = 2L,
  CHART           = 3L,
  COMMENT         = 4L,
  FREEFORM        = 5L,
  GROUP           = 6L,
  EMBEDDED_OLE_OBJECT = 7L,
  FORM_CONTROL    = 8L,
  LINE            = 9L,
  LINKED_OLE_OBJECT = 10L,
  LINKED_PICTURE  = 11L,
  OLE_CONTROL_OBJECT = 12L,
  PICTURE         = 13L,
  PLACEHOLDER     = 14L,
  RTF             = 15L,
  SCRIPT_ANCHOR   = 16L,
  TEXT_BOX        = 17L,
  TEXT_EFFECT     = 18L,
  TABLE           = 19L,
  CANVAS          = 20L,
  DIAGRAM         = 21L,
  INK             = 22L,
  INK_COMMENT     = 23L,
  SMART_ART       = 24L,
  MEDIA           = 25L,
  CONNECTOR       = 10L   # same as LINKED_OLE_OBJECT — use Connector class instead
)

# Re-export as MSO for backward compat / convenience
#' MSO shape type enum (alias for MSO_SHAPE_TYPE)
#' @noRd
#' @export
MSO <- MSO_SHAPE_TYPE


#' MSO_AUTO_SHAPE_TYPE — preset geometry names for autoshapes
#'
#' Each entry is a list with `value` (MSO integer constant) and `prst`
#' (OOXML `prst` attribute string). Pass a member to `SlideShapes$add_shape()`.
#'
#' Alias: `MSO_SHAPE`
#'
#' @noRd
#' @export
MSO_AUTO_SHAPE_TYPE <- list(
  ACTION_BUTTON_BACK_OR_PREVIOUS  = list(value = 129L, prst = "actionButtonBackPrevious"),
  ACTION_BUTTON_BEGINNING         = list(value = 131L, prst = "actionButtonBeginning"),
  ACTION_BUTTON_CUSTOM            = list(value = 125L, prst = "actionButtonBlank"),
  ACTION_BUTTON_DOCUMENT          = list(value = 134L, prst = "actionButtonDocument"),
  ACTION_BUTTON_END               = list(value = 132L, prst = "actionButtonEnd"),
  ACTION_BUTTON_FORWARD_OR_NEXT   = list(value = 130L, prst = "actionButtonForwardNext"),
  ACTION_BUTTON_HELP              = list(value = 127L, prst = "actionButtonHelp"),
  ACTION_BUTTON_HOME              = list(value = 126L, prst = "actionButtonHome"),
  ACTION_BUTTON_INFORMATION       = list(value = 128L, prst = "actionButtonInformation"),
  ACTION_BUTTON_MOVIE             = list(value = 136L, prst = "actionButtonMovie"),
  ACTION_BUTTON_RETURN            = list(value = 133L, prst = "actionButtonReturn"),
  ACTION_BUTTON_SOUND             = list(value = 135L, prst = "actionButtonSound"),
  ARC                             = list(value =  25L, prst = "arc"),
  BALLOON                         = list(value = 137L, prst = "wedgeRoundRectCallout"),
  BENT_ARROW                      = list(value =  41L, prst = "bentArrow"),
  BENT_UP_ARROW                   = list(value =  44L, prst = "bentUpArrow"),
  BEVEL                           = list(value =  15L, prst = "bevel"),
  BLOCK_ARC                       = list(value =  20L, prst = "blockArc"),
  CAN                             = list(value =  13L, prst = "can"),
  CHEVRON                         = list(value =  52L, prst = "chevron"),
  CHORD                           = list(value = 161L, prst = "chord"),
  CIRCULAR_ARROW                  = list(value =  60L, prst = "circularArrow"),
  CLOUD                           = list(value = 179L, prst = "cloud"),
  CLOUD_CALLOUT                   = list(value = 108L, prst = "cloudCallout"),
  CORNER                          = list(value = 162L, prst = "corner"),
  CROSS                           = list(value =  11L, prst = "plus"),
  CUBE                            = list(value =  14L, prst = "cube"),
  CURVED_DOWN_ARROW               = list(value =  48L, prst = "curvedDownArrow"),
  CURVED_LEFT_ARROW               = list(value =  46L, prst = "curvedLeftArrow"),
  CURVED_RIGHT_ARROW              = list(value =  45L, prst = "curvedRightArrow"),
  CURVED_UP_ARROW                 = list(value =  47L, prst = "curvedUpArrow"),
  DECAGON                         = list(value = 144L, prst = "decagon"),
  DIAMOND                         = list(value =   4L, prst = "diamond"),
  DODECAGON                       = list(value = 146L, prst = "dodecagon"),
  DONUT                           = list(value =  18L, prst = "donut"),
  DOUBLE_WAVE                     = list(value = 104L, prst = "doubleWave"),
  DOWN_ARROW                      = list(value =  36L, prst = "downArrow"),
  DOWN_ARROW_CALLOUT              = list(value =  56L, prst = "downArrowCallout"),
  DOWN_RIBBON                     = list(value =  98L, prst = "ribbon"),
  EXPLOSION1                      = list(value =  89L, prst = "irregularSeal1"),
  EXPLOSION2                      = list(value =  90L, prst = "irregularSeal2"),
  FLOWCHART_DECISION              = list(value =  63L, prst = "flowChartDecision"),
  FLOWCHART_PROCESS               = list(value =  61L, prst = "flowChartProcess"),
  FLOWCHART_TERMINATOR            = list(value =  69L, prst = "flowChartTerminator"),
  FOLDED_CORNER                   = list(value =  16L, prst = "foldedCorner"),
  FRAME                           = list(value = 158L, prst = "frame"),
  FUNNEL                          = list(value = 174L, prst = "funnel"),
  GEAR_6                          = list(value = 172L, prst = "gear6"),
  GEAR_9                          = list(value = 173L, prst = "gear9"),
  HEART                           = list(value =  21L, prst = "heart"),
  HEPTAGON                        = list(value = 145L, prst = "heptagon"),
  HEXAGON                         = list(value =  10L, prst = "hexagon"),
  HORIZONTAL_SCROLL               = list(value = 102L, prst = "horizontalScroll"),
  ISOSCELES_TRIANGLE              = list(value =   7L, prst = "triangle"),
  LEFT_ARROW                      = list(value =  34L, prst = "leftArrow"),
  LEFT_ARROW_CALLOUT              = list(value =  54L, prst = "leftArrowCallout"),
  LEFT_BRACE                      = list(value =  31L, prst = "leftBrace"),
  LEFT_BRACKET                    = list(value =  29L, prst = "leftBracket"),
  LEFT_RIGHT_ARROW                = list(value =  37L, prst = "leftRightArrow"),
  LEFT_RIGHT_ARROW_CALLOUT        = list(value =  57L, prst = "leftRightArrowCallout"),
  LEFT_UP_ARROW                   = list(value =  43L, prst = "leftUpArrow"),
  MOON                            = list(value =  24L, prst = "moon"),
  NOTCHED_RIGHT_ARROW             = list(value =  50L, prst = "notchedRightArrow"),
  OCTAGON                         = list(value =   6L, prst = "octagon"),
  OVAL                            = list(value =   9L, prst = "ellipse"),
  OVAL_CALLOUT                    = list(value = 107L, prst = "wedgeEllipseCallout"),
  PARALLELOGRAM                   = list(value =   2L, prst = "parallelogram"),
  PENTAGON                        = list(value =  51L, prst = "homePlate"),
  PLAQUE                          = list(value =  28L, prst = "plaque"),
  QUAD_ARROW                      = list(value =  39L, prst = "quadArrow"),
  QUAD_ARROW_CALLOUT              = list(value =  59L, prst = "quadArrowCallout"),
  RECTANGLE                       = list(value =   1L, prst = "rect"),
  REGULAR_PENTAGON                = list(value =  12L, prst = "pentagon"),
  RIGHT_ARROW                     = list(value =  33L, prst = "rightArrow"),
  RIGHT_ARROW_CALLOUT             = list(value =  53L, prst = "rightArrowCallout"),
  RIGHT_BRACE                     = list(value =  32L, prst = "rightBrace"),
  RIGHT_BRACKET                   = list(value =  30L, prst = "rightBracket"),
  RIGHT_TRIANGLE                  = list(value =   8L, prst = "rtTriangle"),
  ROUNDED_RECTANGLE               = list(value =   5L, prst = "roundRect"),
  ROUNDED_RECTANGULAR_CALLOUT     = list(value = 106L, prst = "wedgeRoundRectCallout"),
  ROUND_1_RECTANGLE               = list(value = 151L, prst = "round1Rect"),
  ROUND_2_DIAG_RECTANGLE          = list(value = 153L, prst = "round2DiagRect"),
  ROUND_2_SAME_RECTANGLE          = list(value = 152L, prst = "round2SameRect"),
  SNIP_1_RECTANGLE                = list(value = 155L, prst = "snip1Rect"),
  SNIP_2_DIAG_RECTANGLE           = list(value = 157L, prst = "snip2DiagRect"),
  SNIP_2_SAME_RECTANGLE           = list(value = 156L, prst = "snip2SameRect"),
  SNIP_ROUND_RECTANGLE            = list(value = 154L, prst = "snipRoundRect"),
  STAR_10_POINT                   = list(value = 149L, prst = "star10"),
  STAR_12_POINT                   = list(value = 150L, prst = "star12"),
  STAR_16_POINT                   = list(value =  94L, prst = "star16"),
  STAR_24_POINT                   = list(value =  95L, prst = "star24"),
  STAR_32_POINT                   = list(value =  96L, prst = "star32"),
  STAR_4_POINT                    = list(value =  91L, prst = "star4"),
  STAR_5_POINT                    = list(value =  92L, prst = "star5"),
  STAR_6_POINT                    = list(value = 147L, prst = "star6"),
  STAR_7_POINT                    = list(value = 148L, prst = "star7"),
  STAR_8_POINT                    = list(value =  93L, prst = "star8"),
  STRIPED_RIGHT_ARROW             = list(value =  49L, prst = "stripedRightArrow"),
  SWOOSH_ARROW                    = list(value = 178L, prst = "swooshArrow"),
  TRAPEZOID                       = list(value =   3L, prst = "trapezoid"),
  UP_ARROW                        = list(value =  35L, prst = "upArrow"),
  UP_ARROW_CALLOUT                = list(value =  55L, prst = "upArrowCallout"),
  UP_DOWN_ARROW                   = list(value =  38L, prst = "upDownArrow"),
  UP_DOWN_ARROW_CALLOUT           = list(value =  58L, prst = "upDownArrowCallout"),
  U_TURN_ARROW                    = list(value =  42L, prst = "uturnArrow"),
  VERTICAL_SCROLL                 = list(value = 103L, prst = "verticalScroll"),
  WAVE                            = list(value = 105L, prst = "wave")
)

#' @noRd
#' @export
MSO_SHAPE <- MSO_AUTO_SHAPE_TYPE


#' PP_PLACEHOLDER — placeholder type string constants
#'
#' String values for the `type` attribute of a `<p:ph>` element.
#'
#' @noRd
#' @export
PP_PLACEHOLDER <- list(
  TITLE         = "title",
  BODY          = "body",
  CENTER_TITLE  = "ctrTitle",
  SUBTITLE      = "subTitle",
  DATE          = "dt",
  FOOTER        = "ftr",
  HEADER        = "hdr",
  SLIDE_NUMBER  = "sldNum",
  OBJECT        = "obj",
  PICTURE       = "pic",
  TABLE         = "tbl",
  CHART         = "chart",
  BITMAP        = "clipArt",
  MEDIA_CLIP    = "media",
  ORG_CHART     = "dgm"
)


#' Connector type constants (MSO_CONNECTOR_TYPE)
#'
#' Members have `$value` (integer) and `$prst` (XML preset name).
#' Alias: `MSO_CONNECTOR`.
#'
#' @noRd
#' @export
MSO_CONNECTOR_TYPE <- list(
  STRAIGHT = list(value = 1L, prst = "line"),
  ELBOW    = list(value = 2L, prst = "bentConnector3"),
  CURVE    = list(value = 3L, prst = "curvedConnector3"),
  MIXED    = list(value = -2L, prst = "")
)

#' MSO connector type enum (alias for MSO_CONNECTOR_TYPE)
#' @noRd
#' @export
MSO_CONNECTOR <- MSO_CONNECTOR_TYPE
