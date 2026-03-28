# Shape-related enumeration constants.
#
# Ported from python-pptx/src/pptx/enum/shapes.py.


#' MSO_SHAPE_TYPE — shape type classification constants
#'
#' Integer constants classifying shapes by type, mirroring MSO_SHAPE_TYPE.
#'
#' @keywords internal
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
#' @export
MSO <- MSO_SHAPE_TYPE


#' PP_PLACEHOLDER — placeholder type string constants
#'
#' String values for the `type` attribute of a `<p:ph>` element.
#'
#' @keywords internal
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
