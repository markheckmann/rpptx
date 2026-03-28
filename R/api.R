# Public API for rpptx.
#
# Ported from python-pptx/src/pptx/api.py. Provides the main
# Presentation() constructor function.

#' Create or open a PowerPoint presentation
#'
#' @param pptx Path to a .pptx file, or NULL to use the default template.
#' @return A Presentation object.
#' @include package.R parts-presentation.R
#' @export
pptx_presentation <- function(pptx = NULL) {
  if (is.null(pptx)) {
    pptx <- .default_pptx_path()
  }
  pkg <- Package_open(pptx)
  prs_part <- pkg$presentation_part

  if (!.is_pptx_package(prs_part)) {
    stop("File is not a PowerPoint (.pptx) file", call. = FALSE)
  }

  prs_part$presentation
}


# Path to the bundled default .pptx template
.default_pptx_path <- function() {
  path <- system.file("templates", "default.pptx", package = "rpptx")
  if (path == "") {
    stop("Default template not found in rpptx package", call. = FALSE)
  }
  path
}


# Check if a PresentationPart has a valid PowerPoint content type
.is_pptx_package <- function(prs_part) {
  ct <- prs_part$content_type
  ct %in% c(CT$PML_PRESENTATION_MAIN, CT$PML_PRES_MACRO_MAIN,
            CT$PML_TEMPLATE_MAIN, CT$PML_SLIDESHOW_MAIN)
}
