# Slide-related OPC part classes.
#
# Ported from python-pptx/src/pptx/parts/slide.py.

# ============================================================================
# BaseSlidePart — common base for all slide part types
# ============================================================================

#' Base class for slide parts (slides, layouts, masters)
#'
#' @include opc-package.R oxml-slide.R
#' @keywords internal
#' @export
BaseSlidePart <- R6::R6Class(
  "BaseSlidePart",
  inherit = XmlPart,

  active = list(
    # Internal name of this slide (from p:cSld/@name)
    name = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      self$element$cSld$name
    }
  )
)


# ============================================================================
# SlidePart — /ppt/slides/slideN.xml
# ============================================================================

#' Slide part
#'
#' @keywords internal
#' @export
SlidePart <- R6::R6Class(
  "SlidePart",
  inherit = BaseSlidePart,

  public = list(
    # Return the Slide domain object (lazily created)
    get_slide = function() {
      if (is.null(private$.slide)) {
        private$.slide <- Slide$new(self$element, self)
      }
      private$.slide
    },

    # True if this slide has a notes slide
    has_notes_slide = function() {
      tryCatch({
        self$part_related_by(RT$NOTES_SLIDE)
        TRUE
      }, error = function(e) FALSE)
    },

    # Create and return a new ChartPart for chart_type/chart_data.
    add_chart_part = function(chart_type, chart_data) {
      ChartPart_new(chart_type, chart_data, self$package)
    }
  ),

  active = list(
    slide = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      self$get_slide()
    },

    slide_id = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      self$package$presentation_part$slide_id(self)
    },

    slide_layout = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      slide_layout_part <- self$part_related_by(RT$SLIDE_LAYOUT)
      slide_layout_part$slide_layout
    }
  ),

  private = list(
    .slide = NULL
  )
)

# Create a new blank SlidePart
SlidePart_new <- function(partname, package, slide_layout_part) {
  slide_element <- new_ct_slide()
  part <- SlidePart$new(partname, CT$PML_SLIDE, package, slide_element)
  part$relate_to(slide_layout_part, RT$SLIDE_LAYOUT)
  part
}


# ============================================================================
# SlideLayoutPart — /ppt/slideLayouts/slideLayoutN.xml
# ============================================================================

#' Slide layout part
#'
#' @keywords internal
#' @export
SlideLayoutPart <- R6::R6Class(
  "SlideLayoutPart",
  inherit = BaseSlidePart,

  public = list(
    # Return SlideLayout related by rId
    related_slide_layout = function(rId) {
      self$related_part(rId)$slide_layout
    }
  ),

  active = list(
    slide_layout = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      if (is.null(private$.slide_layout)) {
        private$.slide_layout <- SlideLayout$new(self$element, self)
      }
      private$.slide_layout
    },

    slide_master = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      self$part_related_by(RT$SLIDE_MASTER)$slide_master
    }
  ),

  private = list(
    .slide_layout = NULL
  )
)


# ============================================================================
# SlideMasterPart — /ppt/slideMasters/slideMasterN.xml
# ============================================================================

#' Slide master part
#'
#' @keywords internal
#' @export
SlideMasterPart <- R6::R6Class(
  "SlideMasterPart",
  inherit = BaseSlidePart,

  public = list(
    # Return SlideLayout related by rId
    related_slide_layout = function(rId) {
      self$related_part(rId)$slide_layout
    }
  ),

  active = list(
    slide_master = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      if (is.null(private$.slide_master)) {
        private$.slide_master <- SlideMaster$new(self$element, self)
      }
      private$.slide_master
    }
  ),

  private = list(
    .slide_master = NULL
  )
)


# ============================================================================
# Register part types
# ============================================================================

.onLoad_parts_slide <- function() {
  register_part_type(CT$PML_SLIDE, SlidePart)
  register_part_type(CT$PML_SLIDE_LAYOUT, SlideLayoutPart)
  register_part_type(CT$PML_SLIDE_MASTER, SlideMasterPart)
}
