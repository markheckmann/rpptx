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
    },

    # Return list(image_part, rId) for image_file; reuses existing part if possible.
    get_or_add_image_part = function(image_file) {
      image_part <- self$package$get_or_add_image_part(image_file)
      # relate_to returns existing rId if image_part is already related
      rId <- self$relate_to(image_part, RT$IMAGE)
      list(image_part = image_part, rId = rId)
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
    },

    # NotesSlide for this slide. Creates a new NotesSlidePart if absent.
    notes_slide = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      notes_slide_part <- tryCatch(
        self$part_related_by(RT$NOTES_SLIDE),
        error = function(e) {
          notes_master_part <- self$package$presentation_part$notes_master_part
          nsp <- NotesSlidePart_new(self$package, self, notes_master_part)
          self$relate_to(nsp, RT$NOTES_SLIDE)
          nsp
        }
      )
      notes_slide_part$notes_slide
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
# NotesMasterPart — /ppt/notesMasters/notesMaster1.xml
# ============================================================================

#' Notes master part
#'
#' @keywords internal
#' @export
NotesMasterPart <- R6::R6Class(
  "NotesMasterPart",
  inherit = BaseSlidePart,

  active = list(
    notes_master = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      if (is.null(private$.notes_master)) {
        private$.notes_master <- NotesMaster$new(self$element, self)
      }
      private$.notes_master
    }
  ),

  private = list(
    .notes_master = NULL
  )
)

# Create a default NotesMasterPart (loads from template, relates to a theme part).
NotesMasterPart_create_default <- function(package) {
  notes_master_elm <- parse_from_template("notesMaster")
  part <- NotesMasterPart$new(
    PackURI("/ppt/notesMasters/notesMaster1.xml"),
    CT$PML_NOTES_MASTER,
    package,
    notes_master_elm
  )
  # Create and relate a default theme part for the notes master
  theme_elm     <- parse_from_template("theme")
  theme_partname <- package$next_partname("/ppt/notesMasters/theme%d.xml")
  theme_part    <- XmlPart$new(theme_partname, CT$OFC_THEME, package, theme_elm)
  part$relate_to(theme_part, RT$THEME)
  part
}


# ============================================================================
# NotesSlidePart — /ppt/notesSlides/notesSlideN.xml
# ============================================================================

#' Notes slide part
#'
#' @keywords internal
#' @export
NotesSlidePart <- R6::R6Class(
  "NotesSlidePart",
  inherit = BaseSlidePart,

  active = list(
    notes_master = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      self$part_related_by(RT$NOTES_MASTER)$notes_master
    },

    notes_slide = function(value) {
      if (!missing(value)) stop("Read-only", call. = FALSE)
      if (is.null(private$.notes_slide)) {
        private$.notes_slide <- NotesSlide$new(self$element, self)
      }
      private$.notes_slide
    }
  ),

  private = list(
    .notes_slide = NULL
  )
)

# Create a new NotesSlidePart for slide_part, using (or creating) notes_master_part.
# Clones cloneable placeholders from the notes master.
NotesSlidePart_new <- function(package, slide_part, notes_master_part) {
  partname  <- package$next_partname("/ppt/notesSlides/notesSlide%d.xml")
  notes_elm <- parse_from_template("notes")
  notes_slide_part <- NotesSlidePart$new(
    partname, CT$PML_NOTES_SLIDE, package, notes_elm
  )
  notes_slide_part$relate_to(notes_master_part, RT$NOTES_MASTER)
  notes_slide_part$relate_to(slide_part, RT$SLIDE)

  # Clone cloneable placeholders from notes master into notes slide
  .clone_notes_master_placeholders(
    notes_slide_part$notes_slide,
    notes_master_part$notes_master
  )

  notes_slide_part
}

# Clone slide-image, body, and slide-number placeholders from notes master.
.clone_notes_master_placeholders <- function(notes_slide, notes_master) {
  cloneable_types <- c("sldImg", "body", "sldNum")
  spTree <- notes_slide$element$spTree

  master_shapes <- notes_master$shapes$to_list()
  for (shape in master_shapes) {
    ph <- tryCatch(shape$element$ph, error = function(e) NULL)
    if (is.null(ph)) next
    ph_type <- tryCatch(ph$type, error = function(e) NULL)
    if (is.null(ph_type)) next
    if (!(ph_type %in% cloneable_types)) next

    # Deep-copy via XML serialization
    node_xml <- xml2::xml_serialize(shape$element$get_node(), NULL)
    cloned   <- xml2::xml_unserialize(node_xml)
    xml2::xml_add_child(spTree$get_node(), xml2::xml_root(cloned))
  }
  invisible(notes_slide)
}


# ============================================================================
# Register part types
# ============================================================================

.onLoad_parts_slide <- function() {
  register_part_type(CT$PML_SLIDE,         SlidePart)
  register_part_type(CT$PML_SLIDE_LAYOUT,  SlideLayoutPart)
  register_part_type(CT$PML_SLIDE_MASTER,  SlideMasterPart)
  register_part_type(CT$PML_NOTES_MASTER,  NotesMasterPart)
  register_part_type(CT$PML_NOTES_SLIDE,   NotesSlidePart)
}
