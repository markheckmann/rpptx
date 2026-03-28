# Package — top-level container for a .pptx file.
#
# Ported from python-pptx/src/pptx/package.py. Extends OpcPackage with
# PowerPoint-specific functionality.

#' PowerPoint package
#'
#' Top-level container for a .pptx file. Extends OpcPackage with
#' presentation-specific behavior.
#'
#' @include opc-package.R parts-coreprops.R
#' @keywords internal
#' @export
Package <- R6::R6Class(
  "Package",
  inherit = OpcPackage,

  active = list(
    # The PresentationPart for this package
    presentation_part = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      self$main_document_part
    },

    # CorePropertiesPart (lazy; creates a default if absent)
    core_properties = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      if (is.null(private$.core_properties)) {
        private$.core_properties <- tryCatch(
          self$part_related_by(RT$CORE_PROPERTIES),
          error = function(e) {
            cp <- CorePropertiesPart_default(self)
            self$relate_to(cp, RT$CORE_PROPERTIES)
            cp
          }
        )
      }
      private$.core_properties
    }
  ),

  private = list(
    .core_properties = NULL
  )
)

#' Open a Package from a .pptx file
#' @param pkg_file Path to a .pptx file.
#' @return A Package instance.
#' @keywords internal
#' @export
Package_open <- function(pkg_file) {
  pkg <- Package$new(pkg_file)
  pkg$.__enclos_env__$private$.load()
  pkg
}
