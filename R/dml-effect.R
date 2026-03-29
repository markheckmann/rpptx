# Visual effects on shapes — shadow.
#
# Ported from python-pptx/src/pptx/dml/effect.py.

# ============================================================================
# ShadowFormat
# ============================================================================

#' Shape shadow effect
#'
#' Provides access to the shadow effect for a shape. Obtain via `shape$shadow`.
#'
#' @include oxml-shapes.R
#' @keywords internal
#' @export
ShadowFormat <- R6::R6Class(
  "ShadowFormat",

  public = list(
    # spPr may also be a grpSpPr; both have an a:effectLst child.
    initialize = function(spPr) private$.spPr <- spPr
  ),

  active = list(
    # TRUE if this shape inherits its shadow from the style hierarchy.
    # Setting to TRUE removes any explicitly-defined effects (restores
    # inheritance). Setting to FALSE establishes an empty effectLst,
    # which suppresses all inherited effects.
    inherit = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) {
          # Remove explicitly-defined effectLst → restore inheritance
          nd <- xml2::xml_find_first(
            private$.spPr$get_node(), "a:effectLst",
            ns = c(a = .nsmap[["a"]])
          )
          if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
        } else {
          # Ensure effectLst is present (even empty) → breaks inheritance
          nd <- xml2::xml_find_first(
            private$.spPr$get_node(), "a:effectLst",
            ns = c(a = .nsmap[["a"]])
          )
          if (inherits(nd, "xml_missing")) {
            xml2::xml_add_child(private$.spPr$get_node(), "a:effectLst",
                                xmlns = .nsmap[["a"]])
          }
        }
        return(invisible(value))
      }
      # Read: TRUE if no explicit effectLst exists
      nd <- xml2::xml_find_first(
        private$.spPr$get_node(), "a:effectLst",
        ns = c(a = .nsmap[["a"]])
      )
      inherits(nd, "xml_missing")
    }
  ),

  private = list(.spPr = NULL)
)
