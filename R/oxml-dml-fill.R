# Custom XML element classes for DML fill elements.
#
# Ported from python-pptx/src/pptx/oxml/dml/fill.py.

# ============================================================================
# CT_NoFill — <a:noFill>
# ============================================================================

#' @include oxml-xmlchemy.R oxml-simpletypes.R oxml-dml-color.R
#' @keywords internal
CT_NoFill <- define_oxml_element(
  classname = "CT_NoFill",
  tag       = "a:noFill"
)

#' @keywords internal
CT_GroupFillProperties <- define_oxml_element(
  classname = "CT_GroupFillProperties",
  tag       = "a:grpFill"
)


# ============================================================================
# CT_GradientStopList — <a:gsLst>
# ============================================================================

#' @keywords internal
CT_GradientStopList <- define_oxml_element(
  classname = "CT_GradientStopList",
  tag       = "a:gsLst",
  children  = list(
    gs = one_or_more("a:gs")
  )
)


# ============================================================================
# CT_LinearShadeProperties — <a:lin>  (gradient linear direction)
# NOTE: "a:lin" here is a:lin inside a:gradFill — NOT a:ln (line properties)
# ============================================================================

#' @keywords internal
CT_LinearShadeProperties <- define_oxml_element(
  classname  = "CT_LinearShadeProperties",
  tag        = "a:lin",
  attributes = list(
    ang    = optional_attribute("ang",    ST_PositiveFixedAngle, default = 0.0),
    scaled = optional_attribute("scaled", XsdBoolean,            default = FALSE)
  )
)


# ============================================================================
# CT_GradientFillProperties — <a:gradFill>
# ============================================================================

#' @keywords internal
CT_GradientFillProperties <- define_oxml_element(
  classname = "CT_GradientFillProperties",
  tag       = "a:gradFill",
  children  = list(
    gsLst = zero_or_one("a:gsLst", successors = c("a:lin", "a:path", "a:tileRect")),
    lin   = zero_or_one("a:lin",   successors = c("a:path", "a:tileRect")),
    path  = zero_or_one("a:path",  successors = c("a:tileRect")),
    tileRect = zero_or_one("a:tileRect")
  )
)


# ============================================================================
# CT_PatternFillProperties — <a:pattFill>
# ============================================================================

#' @keywords internal
CT_PatternFillProperties <- R6::R6Class(
  "CT_PatternFillProperties",
  inherit = BaseOxmlElement,

  active = list(
    # Pattern preset string
    prst = function(value) {
      if (!missing(value)) {
        xml2::xml_set_attr(private$.node, "prst", as.character(value))
        return(invisible(value))
      }
      v <- xml2::xml_attr(private$.node, "prst")
      if (is.na(v)) NULL else v
    },

    # <a:fgClr> child (foreground color container)
    fgClr = function() {
      result <- xml2::xml_find_first(private$.node, "a:fgClr",
                                     ns = c(a = .nsmap[["a"]]))
      if (inherits(result, "xml_missing")) return(NULL)
      wrap_element(result)
    },

    # <a:bgClr> child (background color container)
    bgClr = function() {
      result <- xml2::xml_find_first(private$.node, "a:bgClr",
                                     ns = c(a = .nsmap[["a"]]))
      if (inherits(result, "xml_missing")) return(NULL)
      wrap_element(result)
    }
  ),

  public = list(
    # Return existing <a:fgClr> or create a new one with a default srgbClr
    get_or_add_fgClr = function() {
      fg <- self$fgClr
      if (!is.null(fg)) return(fg)
      a <- .nsmap[["a"]]
      nd <- xml2::read_xml(sprintf(
        '<a:fgClr xmlns:a="%s"><a:srgbClr val="000000"/></a:fgClr>', a))
      xml2::xml_add_child(private$.node, xml2::xml_root(nd), .where = 0L)
      self$fgClr
    },

    # Return existing <a:bgClr> or create a new one with a default srgbClr
    get_or_add_bgClr = function() {
      bg <- self$bgClr
      if (!is.null(bg)) return(bg)
      a <- .nsmap[["a"]]
      nd <- xml2::read_xml(sprintf(
        '<a:bgClr xmlns:a="%s"><a:srgbClr val="FFFFFF"/></a:bgClr>', a))
      xml2::xml_add_child(private$.node, xml2::xml_root(nd))
      self$bgClr
    }
  )
)


# ============================================================================
# CT_BlipFillProperties — <a:blipFill> (picture fill)
# ============================================================================

#' @keywords internal
CT_BlipFillProperties <- define_oxml_element(
  classname = "CT_BlipFillProperties",
  tag       = "a:blipFill",
  children  = list(
    blip     = zero_or_one("a:blip",    successors = c("a:srcRect", "a:tile", "a:stretch")),
    srcRect  = zero_or_one("a:srcRect", successors = c("a:tile", "a:stretch")),
    tile     = zero_or_one("a:tile",    successors = c("a:stretch")),
    stretch  = zero_or_one("a:stretch")
  )
)


# ============================================================================
# Fill-choice tags (used by FillFormat to find/remove fill elements)
# ============================================================================

#' Tags for fill choice elements within spPr/ln (Clark notation)
#' @keywords internal
.fill_choice_tags <- c(
  qn("a:noFill"), qn("a:solidFill"), qn("a:gradFill"),
  qn("a:blipFill"), qn("a:pattFill"), qn("a:grpFill")
)


# ============================================================================
# Helpers to add fill elements (default gradient XML)
# These are used by FillFormat$gradient() to build a default gradient.
# ============================================================================

# Build a default 2-stop linear gradient XML string
.default_gradFill_xml <- function() {
  a <- .nsmap[["a"]]
  sprintf(paste0(
    '<a:gradFill xmlns:a="%s">\n',
    '  <a:gsLst>\n',
    '    <a:gs pos="0">\n',
    '      <a:schemeClr val="accent1"><a:lumMod val="100000"/></a:schemeClr>\n',
    '    </a:gs>\n',
    '    <a:gs pos="100000">\n',
    '      <a:schemeClr val="accent1">\n',
    '        <a:lumMod val="50000"/><a:lumOff val="50000"/>\n',
    '      </a:schemeClr>\n',
    '    </a:gs>\n',
    '  </a:gsLst>\n',
    '  <a:lin ang="5400000" scaled="0"/>\n',
    '</a:gradFill>'
  ), a)
}


# ============================================================================
# Element registration
# ============================================================================

.onLoad_oxml_dml_fill <- function() {
  register_element_cls("a:noFill",   CT_NoFill)
  register_element_cls("a:grpFill",  CT_GroupFillProperties)
  register_element_cls("a:gsLst",    CT_GradientStopList)
  register_element_cls("a:lin",      CT_LinearShadeProperties)
  register_element_cls("a:gradFill", CT_GradientFillProperties)
  register_element_cls("a:pattFill", CT_PatternFillProperties)
  register_element_cls("a:blipFill", CT_BlipFillProperties)
}
