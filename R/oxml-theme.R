# Custom XML element classes for theme-related XML elements.
#
# Ported from python-pptx/src/pptx/oxml/theme.py.

#' @include oxml-xmlchemy.R
#' @noRd


# ============================================================================
# CT_OfficeStyleSheet — <a:theme>
# ============================================================================

#' @noRd
CT_OfficeStyleSheet <- R6::R6Class(
  "CT_OfficeStyleSheet",
  inherit = BaseOxmlElement,

  public = list(
    # Return a new <a:theme> element from the default template.
    new_default = function() {
      parse_from_template("theme")
    }
  ),

  active = list(
    # The theme name attribute, or NULL if absent.
    name = function(value) {
      if (!missing(value)) {
        if (is.null(value)) {
          self$remove_attr("name")
        } else {
          self$set_attr("name", as.character(value))
        }
        return(invisible(value))
      }
      self$get_attr("name")
    }
  )
)


# ============================================================================
# Registration
# ============================================================================

.onLoad_oxml_theme <- function() {
  register_element_cls("a:theme", CT_OfficeStyleSheet)
}
