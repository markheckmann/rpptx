# PresentationPart — the OPC part for /ppt/presentation.xml.
#
# Ported from python-pptx/src/pptx/parts/presentation.py.

#' Presentation Part
#'
#' OPC part for the presentation XML (/ppt/presentation.xml).
#'
#' @include opc-package.R oxml-presentation.R presentation.R
#' @keywords internal
#' @export
PresentationPart <- R6::R6Class(
  "PresentationPart",
  inherit = XmlPart,

  public = list(

    # Return the Presentation domain object for this part
    # Creates it lazily on first access.
    get_presentation = function() {
      if (is.null(private$.presentation)) {
        private$.presentation <- Presentation$new(self$element, self)
      }
      private$.presentation
    },

    # Return the Slide related by rId
    related_slide = function(rId) {
      slide_part <- self$related_part(rId)
      slide_part$get_slide()
    },

    # Return the SlideMaster related by rId
    related_slide_master = function(rId) {
      master_part <- self$related_part(rId)
      master_part$get_slide_master()
    },

    # Save to a file path
    save = function(path) {
      self$package$save(path)
    },

    # Return the slide ID for a given slide part
    slide_id = function(slide_part) {
      sld_id_lst <- self$element$sldIdLst
      if (is.null(sld_id_lst)) {
        stop("no sldIdLst element in presentation", call. = FALSE)
      }
      for (sld_id in sld_id_lst$sldId_lst) {
        rel_part <- self$related_part(sld_id$rId)
        if (identical(rel_part, slide_part)) {
          return(sld_id$id)
        }
      }
      stop("slide_part not found in presentation", call. = FALSE)
    },

    # Rename slide parts to sequential order
    rename_slide_parts = function(rIds) {
      for (i in seq_along(rIds)) {
        slide_part <- self$related_part(rIds[i])
        new_partname <- PackURI(sprintf("/ppt/slides/slide%d.xml", i))
        slide_part$partname <- new_partname
      }
    }
  ),

  active = list(
    # The Presentation domain object
    presentation = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      self$get_presentation()
    }
  ),

  private = list(
    .presentation = NULL
  )
)

# Class method: load from blob
PresentationPart_load <- function(partname, content_type, package, blob) {
  XmlPart_load(PresentationPart, partname, content_type, package, blob)
}

# Register PresentationPart for its content types
.onLoad_parts_presentation <- function() {
  register_part_type(CT$PML_PRESENTATION_MAIN, PresentationPart)
  register_part_type(CT$PML_PRES_MACRO_MAIN, PresentationPart)
  register_part_type(CT$PML_TEMPLATE_MAIN, PresentationPart)
  register_part_type(CT$PML_SLIDESHOW_MAIN, PresentationPart)
}
