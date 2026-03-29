# PresentationPart — the OPC part for /ppt/presentation.xml.
#
# Ported from python-pptx/src/pptx/parts/presentation.py.

#' Presentation Part
#'
#' OPC part for the presentation XML (/ppt/presentation.xml).
#'
#' @include opc-package.R oxml-presentation.R presentation.R parts-slide.R
#' @noRd
PresentationPart <- R6::R6Class(
  "PresentationPart",
  inherit = XmlPart,

  public = list(

    # Add a new slide based on slide_layout; return list(rId, slide)
    add_slide = function(slide_layout) {
      partname <- private$.next_slide_partname()
      slide_layout_part <- slide_layout$part
      slide_part <- SlidePart_new(partname, self$package, slide_layout_part)
      rId <- self$relate_to(slide_part, RT$SLIDE)
      list(rId = rId, slide = slide_part$slide)
    },

    # Return the Presentation domain object (lazy)
    get_presentation = function() {
      if (is.null(private$.presentation)) {
        private$.presentation <- Presentation$new(self$element, self)
      }
      private$.presentation
    },

    # Return the Slide with given slide_id, or NULL
    get_slide = function(slide_id) {
      sld_id_lst <- self$element$sldIdLst
      if (is.null(sld_id_lst)) return(NULL)
      for (sld_id in sld_id_lst$sldId_lst) {
        if (sld_id$id == slide_id) {
          return(self$related_part(sld_id$rId)$slide)
        }
      }
      NULL
    },

    # Return the Slide related by rId
    related_slide = function(rId) {
      self$related_part(rId)$slide
    },

    # Return the SlideMaster related by rId
    related_slide_master = function(rId) {
      self$related_part(rId)$slide_master
    },

    # Save to a file path or connection
    save = function(path) {
      self$package$save(path)
    },

    # Return the slide ID integer for a given SlidePart
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

    # Duplicate slide_part and return list(rId, slide) for the new slide.
    # The new SlidePart has cloned XML and a copy of all source relationships.
    duplicate_slide = function(slide_part) {
      cloned_elm <- rpptx_parse_xml(slide_part$blob)
      partname   <- self$package$next_partname("/ppt/slides/slide%d.xml")
      new_part   <- SlidePart$new(partname, CT$PML_SLIDE, self$package, cloned_elm)
      # Copy relationships in rId-numeric order so the new rIds match the XML
      src_rels <- slide_part$rels$values()
      rId_nums <- vapply(src_rels, function(rel) {
        m <- regmatches(rel$rId, regexec("^rId([0-9]+)$", rel$rId))[[1]]
        if (length(m) < 2L) 0L else as.integer(m[2L])
      }, integer(1L))
      for (rel in src_rels[order(rId_nums)]) {
        if (rel$is_external) {
          new_part$relate_to(rel$target_ref, rel$reltype, is_external = TRUE)
        } else {
          new_part$relate_to(rel$target_part, rel$reltype)
        }
      }
      rId <- self$relate_to(new_part, RT$SLIDE)
      list(rId = rId, slide = new_part$slide)
    },

    # Rename slide parts to /ppt/slides/slide1.xml, slide2.xml, ...
    rename_slide_parts = function(rIds) {
      for (i in seq_along(rIds)) {
        slide_part <- self$related_part(rIds[i])
        slide_part$partname <- PackURI(sprintf("/ppt/slides/slide%d.xml", i))
      }
    }
  ),

  active = list(
    # The Presentation domain object
    presentation = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      self$get_presentation()
    },

    # NotesMaster domain object for this presentation.
    # Creates a default NotesMasterPart + part relationship if absent.
    notes_master_part = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      tryCatch(
        self$part_related_by(RT$NOTES_MASTER),
        error = function(e) {
          part <- NotesMasterPart_create_default(self$package)
          self$relate_to(part, RT$NOTES_MASTER)
          part
        }
      )
    }
  ),

  private = list(
    .presentation = NULL,

    # Next available slide partname: /ppt/slides/slideN.xml
    .next_slide_partname = function() {
      sld_id_lst <- self$element$sldIdLst
      n <- if (is.null(sld_id_lst)) 0L else length(sld_id_lst$sldId_lst)
      self$package$next_partname("/ppt/slides/slide%d.xml")
    }
  )
)

# Register PresentationPart for its content types
.onLoad_parts_presentation <- function() {
  register_part_type(CT$PML_PRESENTATION_MAIN, PresentationPart)
  register_part_type(CT$PML_PRES_MACRO_MAIN, PresentationPart)
  register_part_type(CT$PML_TEMPLATE_MAIN, PresentationPart)
  register_part_type(CT$PML_SLIDESHOW_MAIN, PresentationPart)
}
