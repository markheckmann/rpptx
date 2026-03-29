# Custom XML element classes for DML line/border elements.
#
# Ported from python-pptx/src/pptx/oxml/dml/line.py.

# ============================================================================
# CT_PresetLineDashProperties — <a:prstDash>
# ============================================================================

#' @include oxml-xmlchemy.R oxml-simpletypes.R
#' @keywords internal
CT_PresetLineDashProperties <- define_oxml_element(
  classname  = "CT_PresetLineDashProperties",
  tag        = "a:prstDash",
  attributes = list(
    val = optional_attribute("val", XsdString, default = "solid")
  )
)


# ============================================================================
# CT_LineProperties — <a:ln>
# ============================================================================

#' @keywords internal
CT_LineProperties <- R6::R6Class(
  "CT_LineProperties",
  inherit = BaseOxmlElement,

  public = list(
    # Return <a:prstDash>, creating it if absent
    get_or_add_prstDash = function() {
      existing <- self$prstDash
      if (!is.null(existing)) return(existing)
      xmlns_a <- .nsmap[["a"]]
      node <- xml2::xml_add_child(
        private$.node,
        xml2::read_xml(sprintf('<a:prstDash xmlns:a="%s" val="solid"/>', xmlns_a))
      )
      wrap_element(xml2::xml_child(private$.node, xml2::xml_length(private$.node)))
    },

    # Remove any fill choice child (noFill/solidFill/gradFill/pattFill)
    remove_fill = function() {
      for (child in xml2::xml_children(private$.node)) {
        if (.get_clark_name(child) %in% .fill_choice_tags) {
          xml2::xml_remove(child)
        }
      }
      invisible(NULL)
    },

    # Add <a:noFill/> as fill choice
    get_or_add_noFill = function() {
      self$remove_fill()
      xmlns_a <- .nsmap[["a"]]
      xml2::xml_add_child(
        private$.node,
        xml2::read_xml(sprintf('<a:noFill xmlns:a="%s"/>', xmlns_a))
      )
      wrap_element(xml2::xml_child(private$.node, 1L))  # noFill is first child
    },

    # Add <a:solidFill/> as fill choice; return the solidFill element
    get_or_add_solidFill = function() {
      self$remove_fill()
      xmlns_a <- .nsmap[["a"]]
      xml2::xml_add_child(
        private$.node,
        xml2::read_xml(sprintf('<a:solidFill xmlns:a="%s"/>', xmlns_a))
      )
      wrap_element(xml2::xml_child(private$.node, 1L))
    }
  ),

  active = list(
    # Line width in EMU (read/write). NULL removes the attribute (theme default).
    w = function(value) {
      if (!missing(value)) {
        if (is.null(value)) {
          xml2::xml_set_attr(private$.node, "w", NULL)
        } else {
          xml2::xml_set_attr(private$.node, "w", ST_LineWidth$to_xml(value))
        }
        return(invisible(value))
      }
      w_str <- xml2::xml_attr(private$.node, "w")
      if (is.na(w_str)) return(NULL)
      ST_LineWidth$from_xml(w_str)
    },

    # <a:prstDash> child, or NULL
    prstDash = function() {
      result <- xml2::xml_find_first(private$.node, "a:prstDash",
                                     ns = c(a = .nsmap[["a"]]))
      if (inherits(result, "xml_missing")) return(NULL)
      wrap_element(result)
    },

    # The fill choice child (noFill/solidFill/etc.), or NULL
    eg_fillProperties = function() {
      for (child in xml2::xml_children(private$.node)) {
        if (.get_clark_name(child) %in% .fill_choice_tags) {
          return(wrap_element(child))
        }
      }
      NULL
    }
  )
)


# ============================================================================
# Element registration
# ============================================================================

.onLoad_oxml_dml_line <- function() {
  register_element_cls("a:prstDash", CT_PresetLineDashProperties)
  register_element_cls("a:ln",       CT_LineProperties)
}
