# Slide-related domain objects.
#
# Ported from python-pptx/src/pptx/slide.py.
# Provides Slide, SlideLayout, SlideMaster, Slides, SlideLayouts, SlideMasters.

# ============================================================================
# _BaseSlide — shared base for all slide types
# ============================================================================

#' Base class for slide objects (slides, layouts, masters)
#'
#' @include shared.R
#' @keywords internal
#' @export
.BaseSlide <- R6::R6Class(
  ".BaseSlide",
  inherit = PartElementProxy,

  active = list(
    # Internal name of this slide (read/write)
    name = function(value) {
      if (!missing(value)) {
        self$element$cSld$name <- value %||% ""
        return(invisible(value))
      }
      self$element$cSld$name
    }
  )
)


# ============================================================================
# Slide — a single slide
# ============================================================================

#' Slide object
#'
#' Provides access to slide properties.
#'
#' @keywords internal
#' @export
Slide <- R6::R6Class(
  "Slide",
  inherit = .BaseSlide,

  active = list(
    # True if slide inherits background from master
    follow_master_background = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      is.null(self$element$bg)
    },

    # True if a notes slide exists for this slide
    has_notes_slide = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      self$part$has_notes_slide()
    },

    # Integer slide ID (unique within presentation)
    slide_id = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      self$part$slide_id
    },

    # SlideLayout this slide inherits from
    slide_layout = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      self$part$slide_layout
    },

    # ShapeCollection (stub — shapes implemented in Phase 4)
    shapes = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      if (is.null(private$.shapes)) {
        private$.shapes <- SlideShapes$new(self$element$spTree, self)
      }
      private$.shapes
    },

    # Placeholder collection (stub)
    placeholders = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      NULL
    }
  ),

  private = list(
    .shapes = NULL
  )
)


# ============================================================================
# SlideLayout — a slide layout
# ============================================================================

#' Slide layout object
#'
#' @keywords internal
#' @export
SlideLayout <- R6::R6Class(
  "SlideLayout",
  inherit = .BaseSlide,

  public = list(
    # Generate layout placeholders that should be cloned to a new slide
    iter_cloneable_placeholders = function() {
      list()  # Placeholder implementation — full in Phase 4
    }
  ),

  active = list(
    # SlideShapes collection (stub)
    shapes = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      NULL
    },

    # Placeholder collection (stub)
    placeholders = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      NULL
    },

    # SlideMaster this layout inherits from
    slide_master = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      self$part$slide_master
    },

    # Tuple of slides using this layout
    used_by_slides = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      slides <- self$part$package$presentation_part$presentation$slides
      Filter(function(s) identical(s$slide_layout, self), slides$to_list())
    }
  )
)


# ============================================================================
# SlideMaster — a slide master
# ============================================================================

#' Slide master object
#'
#' @keywords internal
#' @export
SlideMaster <- R6::R6Class(
  "SlideMaster",
  inherit = .BaseSlide,

  active = list(
    # SlideLayouts collection for this master
    slide_layouts = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      if (is.null(private$.slide_layouts)) {
        private$.slide_layouts <- SlideLayouts$new(
          self$element$get_or_add_sldLayoutIdLst(),
          self
        )
      }
      private$.slide_layouts
    },

    # MasterShapes collection (stub)
    shapes = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      NULL
    }
  ),

  private = list(
    .slide_layouts = NULL
  )
)


# ============================================================================
# Slides — sequence of slides in a presentation
# ============================================================================

#' Sequence of slides in a presentation
#'
#' Supports indexed access (1-based), `length()`, and iteration via `to_list()`.
#'
#' @keywords internal
#' @export
Slides <- R6::R6Class(
  "Slides",
  inherit = ParentedElementProxy,

  public = list(
    initialize = function(sldIdLst, prs) {
      super$initialize(sldIdLst, prs)
      private$.sldIdLst <- sldIdLst
    },

    # Return Slide at 1-based index
    get = function(idx) {
      sld_id_lst <- private$.sldIdLst$sldId_lst
      if (idx < 1 || idx > length(sld_id_lst)) {
        stop("slide index out of range", call. = FALSE)
      }
      sld_id <- sld_id_lst[[idx]]
      self$part$related_slide(sld_id$rId)
    },

    # Add a new slide based on slide_layout; return the Slide
    add_slide = function(slide_layout) {
      result <- self$part$add_slide(slide_layout)
      rId <- result$rId
      slide <- result$slide
      # Clone layout placeholders onto the new slide (no-op until Phase 4)
      slide$shapes$clone_layout_placeholders(slide_layout)
      private$.sldIdLst$add_sldId(rId)
      slide
    },

    # Return Slide with given slide_id, or default if not found
    get_by_id = function(slide_id, default = NULL) {
      result <- self$part$get_slide(slide_id)
      if (is.null(result)) default else result
    },

    # Return 1-based index of slide in this collection
    index = function(slide) {
      for (i in seq_along(self$to_list())) {
        if (identical(self$get(i), slide)) return(i)
      }
      stop("slide not in collection", call. = FALSE)
    },

    # Return all slides as a list
    to_list = function() {
      lapply(private$.sldIdLst$sldId_lst, function(sld_id) {
        self$part$related_slide(sld_id$rId)
      })
    }
  ),

  private = list(
    .sldIdLst = NULL
  )
)

# S3 length method
#' @export
length.Slides <- function(x) {
  length(x$.__enclos_env__$private$.sldIdLst$sldId_lst)
}

#' @export
`[[.Slides` <- function(x, i) x$get(i)


# ============================================================================
# SlideLayouts — sequence of slide layouts for a master
# ============================================================================

#' Sequence of slide layouts for a slide master
#'
#' @keywords internal
#' @export
SlideLayouts <- R6::R6Class(
  "SlideLayouts",
  inherit = ParentedElementProxy,

  public = list(
    initialize = function(sldLayoutIdLst, parent) {
      super$initialize(sldLayoutIdLst, parent)
      private$.sldLayoutIdLst <- sldLayoutIdLst
    },

    # Return SlideLayout at 1-based index
    get = function(idx) {
      lst <- private$.sldLayoutIdLst$sldLayoutId_lst
      if (idx < 1 || idx > length(lst)) {
        stop("slide layout index out of range", call. = FALSE)
      }
      entry <- lst[[idx]]
      self$part$related_slide_layout(entry$rId)
    },

    # Return SlideLayout by name, or default
    get_by_name = function(name, default = NULL) {
      for (i in seq_len(length(self))) {
        layout <- self$get(i)
        if (layout$name == name) return(layout)
      }
      default
    },

    # Return 1-based index of slide_layout
    index = function(slide_layout) {
      for (i in seq_len(length(self))) {
        if (identical(self$get(i), slide_layout)) return(i)
      }
      stop("layout not in this SlideLayouts collection", call. = FALSE)
    },

    # Return all layouts as a list
    to_list = function() {
      lapply(seq_len(length(self)), self$get)
    },

    # Remove a slide layout (must not be in use)
    remove = function(slide_layout) {
      if (length(slide_layout$used_by_slides) > 0) {
        stop("cannot remove slide-layout in use by one or more slides",
             call. = FALSE)
      }
      target_idx <- self$index(slide_layout)
      lst <- private$.sldLayoutIdLst$sldLayoutId_lst
      entry <- lst[[target_idx]]
      private$.sldLayoutIdLst$remove_child(entry)
      slide_layout$slide_master$part$drop_rel(entry$rId)
    }
  ),

  private = list(
    .sldLayoutIdLst = NULL
  )
)

#' @export
length.SlideLayouts <- function(x) {
  length(x$.__enclos_env__$private$.sldLayoutIdLst$sldLayoutId_lst)
}

#' @export
`[[.SlideLayouts` <- function(x, i) x$get(i)


# ============================================================================
# SlideMasters — sequence of slide masters
# ============================================================================

#' Sequence of slide masters in a presentation
#'
#' @keywords internal
#' @export
SlideMasters <- R6::R6Class(
  "SlideMasters",
  inherit = ParentedElementProxy,

  public = list(
    initialize = function(sldMasterIdLst, parent) {
      super$initialize(sldMasterIdLst, parent)
      private$.sldMasterIdLst <- sldMasterIdLst
    },

    # Return SlideMaster at 1-based index
    get = function(idx) {
      lst <- private$.sldMasterIdLst$sldMasterId_lst
      if (idx < 1 || idx > length(lst)) {
        stop("slide master index out of range", call. = FALSE)
      }
      entry <- lst[[idx]]
      self$part$related_slide_master(entry$rId)
    },

    # Return all masters as a list
    to_list = function() {
      lapply(seq_len(length(self)), self$get)
    }
  ),

  private = list(
    .sldMasterIdLst = NULL
  )
)

#' @export
length.SlideMasters <- function(x) {
  length(x$.__enclos_env__$private$.sldMasterIdLst$sldMasterId_lst)
}

#' @export
`[[.SlideMasters` <- function(x, i) x$get(i)


# ============================================================================
# SlideShapes — stub shape collection (full implementation in Phase 4)
# ============================================================================

#' Shape collection for a slide (stub)
#'
#' @keywords internal
#' @export
SlideShapes <- R6::R6Class(
  "SlideShapes",
  inherit = ParentedElementProxy,

  public = list(
    initialize = function(spTree, parent) {
      super$initialize(spTree, parent)
      private$.spTree <- spTree
    },

    # No-op until Phase 4 (placeholder cloning)
    clone_layout_placeholders = function(slide_layout) {
      invisible(NULL)
    },

    # Return all shapes as a list (stub — empty for now)
    to_list = function() {
      list()
    }
  ),

  private = list(
    .spTree = NULL
  )
)

#' @export
length.SlideShapes <- function(x) {
  0L
}
