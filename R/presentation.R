# Presentation — user-facing presentation object.
#
# Ported from python-pptx/src/pptx/presentation.py.

#' Presentation object
#'
#' User-facing object for a PowerPoint presentation.
#'
#' @include shared.R slide.R
#' @noRd
Presentation <- R6::R6Class(
  "Presentation",
  inherit = PartElementProxy,

  public = list(
    # Save presentation to a file or connection
    save = function(file) {
      self$part$save(file)
    }
  ),

  active = list(
    # Slide width in EMU (read/write)
    slide_width = function(value) {
      if (!missing(value)) {
        sld_sz <- self$element$get_or_add_sldSz()
        sld_sz$cx <- value
        return(invisible(value))
      }
      sld_sz <- self$element$sldSz
      if (is.null(sld_sz)) return(NULL)
      sld_sz$cx
    },

    # Slide height in EMU (read/write)
    slide_height = function(value) {
      if (!missing(value)) {
        sld_sz <- self$element$get_or_add_sldSz()
        sld_sz$cy <- value
        return(invisible(value))
      }
      sld_sz <- self$element$sldSz
      if (is.null(sld_sz)) return(NULL)
      sld_sz$cy
    },

    # Slides collection (lazy, with rename side-effect on first access)
    slides = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      if (is.null(private$.slides)) {
        sld_id_lst <- self$element$get_or_add_sldIdLst()
        # Rename existing slide parts to canonical order (side-effect)
        rIds <- vapply(sld_id_lst$sldId_lst, function(s) s$rId, character(1))
        if (length(rIds) > 0) self$part$rename_slide_parts(rIds)
        private$.slides <- Slides$new(sld_id_lst, self)
      }
      private$.slides
    },

    # SlideMasters collection (lazy)
    slide_masters = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      if (is.null(private$.slide_masters)) {
        private$.slide_masters <- SlideMasters$new(
          self$element$get_or_add_sldMasterIdLst(),
          self
        )
      }
      private$.slide_masters
    },

    # Convenience: first slide master
    slide_master = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      self$slide_masters$get(1L)
    },

    # Convenience: slide layouts of first slide master
    slide_layouts = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      self$slide_master$slide_layouts
    },

    # CorePropertiesPart for this presentation.
    # No-op setter accepts write-back from chaining: prs$core_properties$title <- "x"
    core_properties = function(value) {
      if (!missing(value)) return(invisible(NULL))
      self$part$package$core_properties
    }
  ),

  private = list(
    .slides = NULL,
    .slide_masters = NULL
  )
)
