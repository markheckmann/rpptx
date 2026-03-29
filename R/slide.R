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

    # Placeholder collection
    placeholders = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      SlidePlaceholders$new(self$element$spTree, self)
    },

    # NotesSlide for this slide (creates one if absent).
    # No-op setter allows chaining: slide$notes_slide$notes_text_frame$text <- "..."
    notes_slide = function(value) {
      if (!missing(value)) return(invisible(NULL))
      self$part$notes_slide
    }
  ),

  private = list(
    .shapes = NULL
  )
)


# ============================================================================
# NotesSlide — proxy for notes slide XML
# ============================================================================

#' Notes slide proxy object
#'
#' Provides access to notes content for a slide, including the notes text frame.
#'
#' @keywords internal
#' @export
NotesSlide <- R6::R6Class(
  "NotesSlide",
  inherit = .BaseSlide,

  active = list(
    # ShapeCollection for shapes on this notes slide
    shapes = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      if (is.null(private$.shapes)) {
        private$.shapes <- SlideShapes$new(self$element$spTree, self)
      }
      private$.shapes
    },

    # The body-type placeholder on this notes slide (contains the notes text), or NULL.
    notes_placeholder = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      phs <- SlidePlaceholders$new(self$element$spTree, self)$to_list()
      for (ph in phs) {
        ph_type <- tryCatch(ph$placeholder_format$type, error = function(e) NULL)
        if (!is.null(ph_type) && ph_type == PP_PLACEHOLDER$BODY) return(ph)
      }
      NULL
    },

    # The TextFrame of the notes placeholder, or NULL if no notes placeholder.
    notes_text_frame = function(value) {
      if (!missing(value)) return(invisible(NULL))
      ph <- self$notes_placeholder
      if (is.null(ph)) return(NULL)
      ph$text_frame
    }
  ),

  private = list(
    .shapes = NULL
  )
)


# ============================================================================
# NotesMaster — proxy for notes master XML
# ============================================================================

#' Notes master proxy object
#'
#' @keywords internal
#' @export
NotesMaster <- R6::R6Class(
  "NotesMaster",
  inherit = .BaseSlide,

  active = list(
    # ShapeCollection for shapes on this notes master
    shapes = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      if (is.null(private$.shapes)) {
        private$.shapes <- SlideShapes$new(self$element$spTree, self)
      }
      private$.shapes
    },

    # Placeholder collection for notes master
    placeholders = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      SlidePlaceholders$new(self$element$spTree, self)
    }
  ),

  private = list(
    .shapes = NULL
  )
)


# ============================================================================
# SlidePlaceholders — placeholder collection for a slide
# ============================================================================

#' Placeholder collection for a slide
#'
#' Provides access to placeholder shapes on a slide. Supports indexed access
#' by placeholder `idx` via `[[`, `length()`, and `to_list()`.
#'
#' @keywords internal
#' @export
SlidePlaceholders <- R6::R6Class(
  "SlidePlaceholders",
  inherit = ParentedElementProxy,

  public = list(
    initialize = function(spTree, parent) {
      super$initialize(spTree, parent)
      private$.spTree <- spTree
    },

    # Return placeholder shape by OOXML idx (0-based integer).
    # Use this to access a placeholder by its OOXML placeholder index.
    get = function(idx) {
      ph_shapes <- private$.ph_shapes()
      for (shape in ph_shapes) {
        if (shape$placeholder_format$idx == idx) return(shape)
      }
      stop(sprintf("no placeholder with idx %d", idx), call. = FALSE)
    },

    # Return placeholder shape at 1-based list position n.
    get_at = function(n) {
      ph_shapes <- private$.ph_shapes()
      if (n < 1L || n > length(ph_shapes)) stop("placeholder position out of range", call. = FALSE)
      ph_shapes[[n]]
    },

    # Return all placeholder shapes as a list
    to_list = function() private$.ph_shapes()
  ),

  private = list(
    .spTree = NULL,

    .ph_shapes = function() {
      if (is.null(private$.spTree)) return(list())
      all_shapes <- lapply(private$.spTree$iter_shape_elms(),
                           function(e) shape_factory(e, self$parent))
      Filter(function(s) isTRUE(s$is_placeholder), all_shapes)
    }
  )
)

#' @export
length.SlidePlaceholders <- function(x) {
  length(x$to_list())
}

#' @export
`[[.SlidePlaceholders` <- function(x, i) x$get_at(i)


# ============================================================================
# SlideMasterPlaceholders — placeholder collection for a slide master
# ============================================================================

#' Placeholder collection for a slide master
#'
#' Like SlidePlaceholders but `get()` accepts a placeholder type string
#' (e.g. "title", "body") rather than an integer idx.
#'
#' @keywords internal
#' @export
SlideMasterPlaceholders <- R6::R6Class(
  "SlideMasterPlaceholders",
  inherit = SlidePlaceholders,

  public = list(
    # Return master placeholder by ph_type string, or NULL if not found.
    get = function(ph_type) {
      ph_shapes <- private$.ph_shapes()
      for (shape in ph_shapes) {
        type <- tryCatch(shape$placeholder_format$type, error = function(e) NULL)
        if (!is.null(type) && type == ph_type) return(shape)
      }
      NULL
    }
  )
)

#' @export
length.SlideMasterPlaceholders <- function(x) length(x$to_list())

#' @export
`[[.SlideMasterPlaceholders` <- function(x, i) x$get_at(i)


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
    # Generate layout placeholder shape elements that should be cloned to a new slide.
    # Excludes latent types: date (dt), footer (ftr), slide number (sldNum).
    iter_cloneable_placeholders = function() {
      latent <- c(PP_PLACEHOLDER$DATE, PP_PLACEHOLDER$FOOTER, PP_PLACEHOLDER$SLIDE_NUMBER)
      spTree <- self$element$spTree
      if (is.null(spTree)) return(list())
      ph_elms <- spTree$iter_ph_elms()
      Filter(function(e) !(e$ph_type %in% latent), ph_elms)
    }
  ),

  active = list(
    # SlideShapes collection (stub)
    shapes = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      NULL
    },

    # SlidePlaceholders collection for this layout
    placeholders = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      spTree <- self$element$spTree
      if (is.null(spTree)) return(SlidePlaceholders$new(NULL, self))
      SlidePlaceholders$new(spTree, self)
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
    },

    # SlideMasterPlaceholders collection for this master
    placeholders = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      spTree <- self$element$spTree
      if (is.null(spTree)) return(SlideMasterPlaceholders$new(NULL, self))
      SlideMasterPlaceholders$new(spTree, self)
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
    },

    # Remove a slide from the presentation.
    # The slide part is unlinked and its relationship dropped. Use with care.
    delete = function(slide) {
      sld_id_lst <- private$.sldIdLst$sldId_lst
      for (sld_id in sld_id_lst) {
        if (identical(self$part$related_slide(sld_id$rId), slide)) {
          rId <- sld_id$rId
          private$.sldIdLst$remove_child(sld_id)
          self$part$drop_rel(rId)
          return(invisible(NULL))
        }
      }
      stop("slide not found in this presentation", call. = FALSE)
    },

    # Move slide to 1-based position idx.
    # The slide is removed from its current position and inserted at idx.
    move = function(slide, idx) {
      sld_id_lst <- private$.sldIdLst$sldId_lst
      n <- length(sld_id_lst)
      if (idx < 1L || idx > n) stop("slide index out of range", call. = FALSE)

      # Find current position
      from_idx <- NULL
      for (i in seq_len(n)) {
        if (identical(self$part$related_slide(sld_id_lst[[i]]$rId), slide)) {
          from_idx <- i
          break
        }
      }
      if (is.null(from_idx)) stop("slide not found in this presentation", call. = FALSE)
      if (from_idx == idx) return(invisible(NULL))

      # Collect (id, rId) pairs in current order, then reorder
      pairs <- lapply(sld_id_lst, function(s) list(id = s$id, rId = s$rId))
      moved  <- pairs[[from_idx]]
      pairs  <- pairs[-from_idx]
      pairs  <- append(pairs, list(moved), after = idx - 1L)

      # Remove all sldId children
      sldIdLst_node <- private$.sldIdLst$get_node()
      for (sld_id in private$.sldIdLst$sldId_lst) {
        xml2::xml_remove(sld_id$get_node())
      }

      # Re-add in new order preserving original IDs
      for (pair in pairs) {
        private$.sldIdLst$`_add_sldId`(id = pair$id, rId = pair$rId)
      }
      invisible(NULL)
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
# SlideShapes — shape collection for a slide
# ============================================================================

#' Shape collection for a slide
#'
#' Supports indexed access (1-based), `length()`, and iteration via `to_list()`.
#'
#' @include shapes-base.R
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

    # Return shape at 1-based index
    get = function(idx) {
      elms <- private$.spTree$iter_shape_elms()
      if (idx < 1L || idx > length(elms)) {
        stop("shape index out of range", call. = FALSE)
      }
      shape_factory(elms[[idx]], self)
    },

    # Return all shapes as a list
    to_list = function() {
      lapply(private$.spTree$iter_shape_elms(),
             function(e) shape_factory(e, self))
    },

    # Add a text box shape at the specified position/size; returns Shape
    add_textbox = function(left, top, width, height) {
      id   <- private$.spTree$next_shape_id()
      name <- sprintf("TextBox %d", id - 1L)
      sp   <- private$.spTree$add_textbox(id, name, left, top, width, height)
      shape_factory(sp, self)
    },

    # Add an autoshape of the given type at position/size; returns Shape.
    # auto_shape_type must be a member of MSO_AUTO_SHAPE_TYPE.
    add_shape = function(auto_shape_type, left, top, width, height) {
      prst <- if (is.list(auto_shape_type)) auto_shape_type[["prst"]] else NULL
      if (is.null(prst)) stop("auto_shape_type must be an MSO_AUTO_SHAPE_TYPE member", call. = FALSE)
      id   <- private$.spTree$next_shape_id()
      name <- sprintf("AutoShape %d", id - 1L)
      sp   <- private$.spTree$add_autoshape(id, name, prst, left, top, width, height)
      shape_factory(sp, self)
    },

    # Add a table at the specified position/size; returns GraphicFrame
    add_table = function(rows, cols, left, top, width, height) {
      id   <- private$.spTree$next_shape_id()
      name <- sprintf("Table %d", id - 1L)
      gf   <- private$.spTree$add_table(id, name, rows, cols,
                                         left, top, width, height)
      shape_factory(gf, self)
    },

    # Add a chart at the specified position/size; returns GraphicFrame.
    # chart_type must be a member of XL_CHART_TYPE.
    # chart_data must be a CategoryChartData, XyChartData, or BubbleChartData.
    add_chart = function(chart_type, left, top, width, height, chart_data) {
      chart_part <- self$part$add_chart_part(chart_type, chart_data)
      rId  <- self$part$relate_to(chart_part, RT$CHART)
      id   <- private$.spTree$next_shape_id()
      name <- sprintf("Chart %d", id - 1L)
      gf   <- private$.spTree$add_chart(id, name, rId, left, top, width, height)
      shape_factory(gf, self)
    },

    # Add a connector shape.
    # connector_type: a member of MSO_CONNECTOR_TYPE (e.g. MSO_CONNECTOR_TYPE$STRAIGHT).
    # begin_x/y and end_x/y are in EMU. flipH/flipV flip the connector geometry.
    add_connector = function(connector_type, begin_x, begin_y, end_x, end_y) {
      prst  <- connector_type$prst
      x     <- min(begin_x, end_x)
      y     <- min(begin_y, end_y)
      cx    <- abs(end_x - begin_x)
      cy    <- abs(end_y - begin_y)
      flipH <- end_x < begin_x
      flipV <- end_y < begin_y
      id    <- private$.spTree$next_shape_id()
      name  <- sprintf("Connector: %s %d", prst, id - 1L)
      cxnSp <- private$.spTree$add_cxnSp(id, name, prst, x, y, cx, cy, flipH, flipV)
      shape_factory(cxnSp, self)
    },

    # Add a picture from a file path or connection.
    # width/height are in EMU (use Inches(), Pt(), etc.); NULL preserves aspect ratio.
    add_picture = function(image_file, left, top, width = NULL, height = NULL) {
      img_info   <- self$part$get_or_add_image_part(image_file)
      image_part <- img_info$image_part
      rId        <- img_info$rId
      dims       <- image_part$scale(width, height)
      cx         <- dims[1]; cy <- dims[2]
      id         <- private$.spTree$next_shape_id()
      name       <- sprintf("Picture %d", id - 1L)
      desc       <- image_part$desc
      pic        <- private$.spTree$add_pic(id, name, desc, rId, left, top, cx, cy)
      shape_factory(pic, self)
    },

    # Group existing shapes into a p:grpSp. shapes_list is a list of shape objects.
    # Each shape is removed from the slide and placed into the new group.
    # Returns a GroupShape wrapping the new p:grpSp element.
    add_group_shape = function(shapes_list) {
      shape_elms <- lapply(shapes_list, function(s) s$.__enclos_env__$private$.element)
      id   <- private$.spTree$next_shape_id()
      name <- sprintf("Group %d", id - 1L)
      grpSp <- private$.spTree$add_grpSp(id, name, shape_elms)
      shape_factory(grpSp, self)
    },

    # Return a FreeformBuilder to specify a freeform shape.
    # start_x / start_y: initial pen position in local coordinates (default 0).
    # scale: local-to-EMU scale; either a single number or c(x_scale, y_scale).
    # Returns a FreeformBuilder object.
    build_freeform = function(start_x = 0, start_y = 0, scale = 1.0) {
      if (length(scale) == 2L) {
        x_scale <- scale[[1L]]; y_scale <- scale[[2L]]
      } else {
        x_scale <- scale; y_scale <- scale
      }
      FreeformBuilder$new(private$.spTree, self, start_x, start_y, x_scale, y_scale)
    },

    # Clone layout placeholders onto this slide.
    # Excluded: date (dt), footer (ftr), slide number (sldNum) — latent placeholders.
    clone_layout_placeholders = function(slide_layout) {
      ph_elms <- slide_layout$iter_cloneable_placeholders()
      for (layout_ph_elm in ph_elms) {
        ph_type  <- layout_ph_elm$ph_type
        ph_orient <- layout_ph_elm$ph_orient
        ph_sz    <- layout_ph_elm$ph_sz
        ph_idx   <- layout_ph_elm$ph_idx
        id       <- private$.spTree$next_shape_id()
        name     <- .ph_basename(ph_type, ph_orient)
        private$.spTree$add_placeholder(id, name, ph_type, ph_orient, ph_sz, ph_idx)
      }
      invisible(NULL)
    }
  ),

  private = list(
    .spTree = NULL
  )
)

# Map placeholder type to base display name
.ph_basename <- function(ph_type, orient = "horz") {
  basenames <- list(
    clipArt = "ClipArt Placeholder",
    body    = "Text Placeholder",
    ctrTitle = "Title",
    chart   = "Chart Placeholder",
    dt      = "Date Placeholder",
    ftr     = "Footer Placeholder",
    hdr     = "Header Placeholder",
    media   = "Media Placeholder",
    obj     = "Content Placeholder",
    dgm     = "SmartArt Placeholder",
    pic     = "Picture Placeholder",
    sldNum  = "Slide Number Placeholder",
    subTitle = "Subtitle",
    tbl     = "Table Placeholder",
    title   = "Title"
  )
  name <- basenames[[ph_type]] %||% sprintf("Placeholder %s", ph_type)
  if (identical(orient, "vert")) name <- paste("Vertical", name)
  name
}

#' @export
length.SlideShapes <- function(x) {
  length(x$.__enclos_env__$private$.spTree$iter_shape_elms())
}

#' @export
`[[.SlideShapes` <- function(x, i) x$get(i)
