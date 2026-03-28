# Custom XML element classes for DML color elements.
#
# Ported from python-pptx/src/pptx/oxml/dml/color.py.

# ============================================================================
# Percentage modifier elements — <a:lumMod>, <a:lumOff>, <a:alpha>
# ============================================================================

#' @include oxml-xmlchemy.R oxml-simpletypes.R
#' @keywords internal
CT_Percentage <- define_oxml_element(
  classname  = "CT_Percentage",
  tag        = "a:lumMod",
  attributes = list(
    val = required_attribute("val", ST_Percentage)
  )
)

#' @keywords internal
CT_LumOff <- define_oxml_element(
  classname  = "CT_LumOff",
  tag        = "a:lumOff",
  attributes = list(
    val = required_attribute("val", ST_Percentage)
  )
)

#' @keywords internal
CT_Alpha <- define_oxml_element(
  classname  = "CT_Alpha",
  tag        = "a:alpha",
  attributes = list(
    val = required_attribute("val", ST_PositiveFixedPercentage)
  )
)


# ============================================================================
# CT_SRgbColor — <a:srgbClr>
# ============================================================================

#' @keywords internal
CT_SRgbColor <- define_oxml_element(
  classname  = "CT_SRgbColor",
  tag        = "a:srgbClr",
  attributes = list(
    val = required_attribute("val", ST_HexColorRGB)
  )
)


# ============================================================================
# CT_SchemeColor — <a:schemeClr>
# ============================================================================

#' @keywords internal
CT_SchemeColor <- define_oxml_element(
  classname  = "CT_SchemeColor",
  tag        = "a:schemeClr",
  children   = list(
    lumMod = zero_or_one("a:lumMod", successors = c("a:lumOff", "a:alpha", "a:shade", "a:tint")),
    lumOff = zero_or_one("a:lumOff", successors = c("a:alpha", "a:shade", "a:tint")),
    alpha  = zero_or_one("a:alpha",  successors = c("a:shade", "a:tint"))
  ),
  attributes = list(
    val = required_attribute("val", XsdString)
  )
)


# ============================================================================
# CT_HslColor — <a:hslClr>
# ============================================================================

#' @keywords internal
CT_HslColor <- define_oxml_element(
  classname  = "CT_HslColor",
  tag        = "a:hslClr",
  attributes = list(
    hue = required_attribute("hue", XsdUnsignedInt),
    sat = required_attribute("sat", ST_Percentage),
    lum = required_attribute("lum", ST_Percentage)
  )
)


# ============================================================================
# CT_ScRgbColor — <a:scrgbClr>
# ============================================================================

#' @keywords internal
CT_ScRgbColor <- define_oxml_element(
  classname  = "CT_ScRgbColor",
  tag        = "a:scrgbClr",
  attributes = list(
    r = required_attribute("r", ST_Percentage),
    g = required_attribute("g", ST_Percentage),
    b = required_attribute("b", ST_Percentage)
  )
)


# ============================================================================
# CT_PresetColor — <a:prstClr>
# ============================================================================

#' @keywords internal
CT_PresetColor <- define_oxml_element(
  classname  = "CT_PresetColor",
  tag        = "a:prstClr",
  attributes = list(
    val = required_attribute("val", XsdString)
  )
)


# ============================================================================
# CT_SystemColor — <a:sysClr>
# ============================================================================

#' @keywords internal
CT_SystemColor <- define_oxml_element(
  classname  = "CT_SystemColor",
  tag        = "a:sysClr",
  attributes = list(
    val      = required_attribute("val",      XsdString),
    lastClr  = optional_attribute("lastClr",  ST_HexColorRGB, default = NULL)
  )
)


# ============================================================================
# Color-choice parent elements — <a:solidFill>, <a:fgClr>, <a:bgClr>, <a:gs>
# These elements contain exactly one "color choice" child (eg_colorChoice).
# We implement them as BaseOxmlElement subclasses with a helper active binding.
# ============================================================================

# Tags that represent a color choice child
.color_choice_tags <- c(
  qn("a:srgbClr"), qn("a:schemeClr"), qn("a:hslClr"),
  qn("a:scrgbClr"), qn("a:prstClr"), qn("a:sysClr")
)

# Mixin-style R6 class providing color-choice access on element wrappers
ColorChoiceParent <- R6::R6Class(
  "ColorChoiceParent",
  inherit = BaseOxmlElement,

  public = list(
    # Remove any existing color choice child
    clear_color_choice = function() {
      for (child in xml2::xml_children(private$.node)) {
        if (.get_clark_name(child) %in% .color_choice_tags) {
          xml2::xml_remove(child)
        }
      }
      invisible(NULL)
    },

    # Add an <a:srgbClr val="RRGGBB"/> child; return wrapped element
    get_or_add_srgbClr = function(rgb_str) {
      self$clear_color_choice()
      xmlns_a <- .nsmap[["a"]]
      node <- xml2::xml_add_child(private$.node,
                                  xml2::read_xml(sprintf(
                                    '<a:srgbClr xmlns:a="%s" val="%s"/>',
                                    xmlns_a, toupper(rgb_str)
                                  )))
      wrap_element(xml2::xml_child(private$.node, xml2::xml_length(private$.node)))
    },

    # Add an <a:schemeClr val="..."/> child; return wrapped element
    get_or_add_schemeClr = function(theme_val) {
      self$clear_color_choice()
      xmlns_a <- .nsmap[["a"]]
      node <- xml2::xml_add_child(private$.node,
                                  xml2::read_xml(sprintf(
                                    '<a:schemeClr xmlns:a="%s" val="%s"/>',
                                    xmlns_a, theme_val
                                  )))
      wrap_element(xml2::xml_child(private$.node, xml2::xml_length(private$.node)))
    }
  ),

  active = list(
    # The color choice child element, or NULL
    eg_colorChoice = function() {
      for (child in xml2::xml_children(private$.node)) {
        if (.get_clark_name(child) %in% .color_choice_tags) {
          return(wrap_element(child))
        }
      }
      NULL
    },

    # Convenience: the <a:srgbClr> child, or NULL
    srgbClr = function() {
      result <- xml2::xml_find_first(private$.node, "a:srgbClr",
                                     ns = c(a = .nsmap[["a"]]))
      if (inherits(result, "xml_missing")) return(NULL)
      wrap_element(result)
    },

    # Convenience: the <a:schemeClr> child, or NULL
    schemeClr = function() {
      result <- xml2::xml_find_first(private$.node, "a:schemeClr",
                                     ns = c(a = .nsmap[["a"]]))
      if (inherits(result, "xml_missing")) return(NULL)
      wrap_element(result)
    }
  )
)


# ============================================================================
# Concrete color-choice-parent elements
# ============================================================================

#' @keywords internal
CT_SolidColorFillProperties <- R6::R6Class(
  "CT_SolidColorFillProperties",
  inherit = ColorChoiceParent
)

#' @keywords internal
CT_ForegroundColor <- R6::R6Class(
  "CT_ForegroundColor",
  inherit = ColorChoiceParent
)

#' @keywords internal
CT_BackgroundColor <- R6::R6Class(
  "CT_BackgroundColor",
  inherit = ColorChoiceParent
)

#' @keywords internal
CT_GradientStop <- R6::R6Class(
  "CT_GradientStop",
  inherit = ColorChoiceParent,

  active = list(
    # Position 0.0–1.0
    position = function(value) {
      if (!missing(value)) {
        xml2::xml_set_attr(private$.node, "pos",
                           ST_PositiveFixedPercentage$to_xml(value))
        return(invisible(value))
      }
      pos_str <- xml2::xml_attr(private$.node, "pos")
      if (is.na(pos_str)) return(NULL)
      ST_PositiveFixedPercentage$from_xml(pos_str)
    }
  )
)


# ============================================================================
# Element registration
# ============================================================================

.onLoad_oxml_dml_color <- function() {
  register_element_cls("a:lumMod",   CT_Percentage)
  register_element_cls("a:lumOff",   CT_LumOff)
  register_element_cls("a:alpha",    CT_Alpha)
  register_element_cls("a:srgbClr",  CT_SRgbColor)
  register_element_cls("a:schemeClr",CT_SchemeColor)
  register_element_cls("a:hslClr",   CT_HslColor)
  register_element_cls("a:scrgbClr", CT_ScRgbColor)
  register_element_cls("a:prstClr",  CT_PresetColor)
  register_element_cls("a:sysClr",   CT_SystemColor)
  register_element_cls("a:solidFill",CT_SolidColorFillProperties)
  register_element_cls("a:fgClr",    CT_ForegroundColor)
  register_element_cls("a:bgClr",    CT_BackgroundColor)
  register_element_cls("a:gs",       CT_GradientStop)
}
