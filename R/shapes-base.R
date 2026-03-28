# Shape domain objects — BaseShape and concrete subclasses.
#
# Ported from python-pptx/src/pptx/shapes/base.py and subclasses.

# ============================================================================
# shape_factory — dispatch to the correct shape proxy class
# ============================================================================

#' Return appropriate shape proxy for a shape XML element
#'
#' Dispatches on element tag: p:sp → Shape, p:pic → Picture, p:cxnSp → Connector,
#' p:grpSp → GroupShape, p:graphicFrame → GraphicFrame. Falls back to BaseShape.
#'
#' @param shape_elm A BaseShapeElement (or subclass) R6 wrapper.
#' @param parent The parent ProvidesPart object (e.g. a Slide).
#' @return An R6 shape proxy.
#' @include enum-shapes.R oxml-shapes.R
#' @keywords internal
#' @export
shape_factory <- function(shape_elm, parent) {
  # Dispatch on registered element class (set by wrap_element / .onLoad_oxml_shapes)
  if (inherits(shape_elm, "CT_Picture"))              return(Picture$new(shape_elm, parent))
  if (inherits(shape_elm, "CT_Connector"))            return(Connector$new(shape_elm, parent))
  if (inherits(shape_elm, "CT_GroupShape"))           return(GroupShape$new(shape_elm, parent))
  if (inherits(shape_elm, "CT_Shape"))                return(Shape$new(shape_elm, parent))
  if (inherits(shape_elm, "CT_GraphicalObjectFrame")) return(GraphicFrame$new(shape_elm, parent))
  BaseShape$new(shape_elm, parent)
}


# ============================================================================
# PlaceholderFormat — wraps <p:ph> element
# ============================================================================

#' Placeholder format descriptor
#'
#' Provides access to placeholder properties such as `idx` and `type`.
#' Accessed via `shape$placeholder_format` on a placeholder shape.
#'
#' @keywords internal
#' @export
PlaceholderFormat <- R6::R6Class(
  "PlaceholderFormat",

  public = list(
    initialize = function(ph_elm) {
      private$.ph <- ph_elm
    },

    # The underlying p:ph element
    element = function() private$.ph
  ),

  active = list(
    # Integer placeholder index
    idx = function() private$.ph$idx,

    # Placeholder type string (e.g. "title", "body", "obj")
    type = function() private$.ph$type
  ),

  private = list(
    .ph = NULL
  )
)


# ============================================================================
# BaseShape — common base for all shape proxy objects
# ============================================================================

#' Base class for shape proxy objects
#'
#' Subclasses include Shape, Picture, Connector, GroupShape, GraphicFrame.
#'
#' @keywords internal
#' @export
BaseShape <- R6::R6Class(
  "BaseShape",

  public = list(
    initialize = function(shape_elm, parent) {
      private$.element <- shape_elm
      private$.parent  <- parent
    }
  ),

  active = list(
    # The underlying XML element (BaseShapeElement subclass)
    element = function() private$.element,

    # The owning slide part
    part = function() private$.parent$part,

    # Integer unique shape ID within slide
    shape_id = function() private$.element$shape_id,

    # Shape display name (read/write)
    name = function(value) {
      if (!missing(value)) {
        private$.element$shape_name <- value
        return(invisible(value))
      }
      private$.element$shape_name
    },

    # Distance from slide left edge in EMU (read/write)
    left = function(value) {
      if (!missing(value)) { private$.element$x <- value; return(invisible(value)) }
      private$.element$x
    },

    # Distance from slide top edge in EMU (read/write)
    top = function(value) {
      if (!missing(value)) { private$.element$y <- value; return(invisible(value)) }
      private$.element$y
    },

    # Shape width in EMU (read/write)
    width = function(value) {
      if (!missing(value)) { private$.element$cx <- value; return(invisible(value)) }
      private$.element$cx
    },

    # Shape height in EMU (read/write)
    height = function(value) {
      if (!missing(value)) { private$.element$cy <- value; return(invisible(value)) }
      private$.element$cy
    },

    # Clockwise rotation in degrees (read/write)
    rotation = function(value) {
      if (!missing(value)) { private$.element$rot <- value; return(invisible(value)) }
      private$.element$rot
    },

    # TRUE only for Shape (p:sp) — overridden there
    has_text_frame = function() FALSE,

    # TRUE if this shape has a <p:ph> element
    is_placeholder = function() private$.element$has_ph_elm,

    # PlaceholderFormat for placeholder shapes; error otherwise
    placeholder_format = function() {
      ph <- private$.element$ph
      if (is.null(ph)) stop("shape is not a placeholder", call. = FALSE)
      PlaceholderFormat$new(ph)
    },

    # FillFormat for this shape's fill properties.
    # No-op setter enables chaining: shape$fill$solid(); shape$fill$fore_color$rgb <- ...
    fill = function(value) {
      if (!missing(value)) return(invisible(NULL))
      spPr <- private$.element$spPr
      if (is.null(spPr)) stop("shape has no spPr element", call. = FALSE)
      FillFormat$new(spPr)
    },

    # LineFormat for this shape's border/line properties.
    # No-op setter enables chaining: shape$line$color$rgb <- ...
    line = function(value) {
      if (!missing(value)) return(invisible(NULL))
      spPr <- private$.element$spPr
      if (is.null(spPr)) stop("shape has no spPr element", call. = FALSE)
      LineFormat$new(spPr)
    },

    # MSO_SHAPE_TYPE member — must be implemented by subclasses
    shape_type = function() {
      stop(paste(class(self)[1], "does not implement shape_type"), call. = FALSE)
    }
  ),

  private = list(
    .element = NULL,
    .parent  = NULL
  )
)


# ============================================================================
# Shape — p:sp (autoshape, textbox, placeholder)
# ============================================================================

#' Shape proxy for p:sp elements
#'
#' Covers autoshapes, text boxes, and placeholders.
#'
#' @keywords internal
#' @export
Shape <- R6::R6Class(
  "Shape",
  inherit = BaseShape,

  active = list(
    has_text_frame = function() TRUE,

    # TextFrame wrapping the shape's p:txBody (created on demand).
    # No-op setter accepts the R6 write-back from x$text_frame$prop <- val.
    text_frame = function(value) {
      if (!missing(value)) return(invisible(NULL))
      txBody <- private$.element$get_or_add_txBody()
      TextFrame$new(txBody, self)
    },

    # Convenience shortcut for all text in this shape (read/write).
    text = function(value) {
      if (!missing(value)) {
        tf <- self$text_frame
        tf$text <- value
        return(invisible(value))
      }
      self$text_frame$text
    },

    shape_type = function() {
      if (self$is_placeholder) return(MSO_SHAPE_TYPE$PLACEHOLDER)
      # Detect textbox via p:cNvSpPr/@txBox="1"
      nodes <- private$.element$xpath("./*[1]/p:cNvSpPr")
      if (length(nodes) > 0) {
        txBox <- xml2::xml_attr(nodes[[1]], "txBox")
        if (!is.na(txBox) && txBox == "1") return(MSO_SHAPE_TYPE$TEXT_BOX)
      }
      MSO_SHAPE_TYPE$AUTO_SHAPE
    }
  )
)


# ============================================================================
# Picture — p:pic
# ============================================================================

#' Shape proxy for p:pic elements
#'
#' @keywords internal
#' @export
Picture <- R6::R6Class(
  "Picture",
  inherit = BaseShape,

  active = list(
    shape_type = function() MSO_SHAPE_TYPE$PICTURE
  )
)


# ============================================================================
# Connector — p:cxnSp
# ============================================================================

#' Shape proxy for p:cxnSp elements
#'
#' @keywords internal
#' @export
Connector <- R6::R6Class(
  "Connector",
  inherit = BaseShape,

  active = list(
    shape_type = function() MSO_SHAPE_TYPE$CONNECTOR
  )
)


# ============================================================================
# GroupShape — p:grpSp
# ============================================================================

#' Shape proxy for p:grpSp (group shape) elements
#'
#' @keywords internal
#' @export
GroupShape <- R6::R6Class(
  "GroupShape",
  inherit = BaseShape,

  active = list(
    shape_type = function() MSO_SHAPE_TYPE$GROUP,

    # Shapes contained within this group (SlideShapes-like collection)
    shapes = function() {
      if (is.null(private$.shapes_cache)) {
        private$.shapes_cache <- GroupShapes$new(private$.element, self)
      }
      private$.shapes_cache
    }
  ),

  private = list(
    .shapes_cache = NULL
  )
)


# ============================================================================
# GraphicFrame — p:graphicFrame (chart, table, OLE)
# ============================================================================

#' Shape proxy for p:graphicFrame elements
#'
#' @keywords internal
#' @export
GraphicFrame <- R6::R6Class(
  "GraphicFrame",
  inherit = BaseShape,

  active = list(
    has_chart = function() {
      grepl("chart", self$.graphic_data_uri, ignore.case = TRUE)
    },

    has_table = function() {
      grepl("table", self$.graphic_data_uri, ignore.case = TRUE)
    },

    # Table object for table-containing graphic frames.
    # No-op setter handles write-back from chaining.
    table = function(value) {
      if (!missing(value)) return(invisible(NULL))
      tbl_elm <- private$.element$tbl
      if (is.null(tbl_elm)) stop("this shape does not contain a table", call. = FALSE)
      Table$new(tbl_elm, self)
    },

    shape_type = function() {
      if (self$has_chart) return(MSO_SHAPE_TYPE$CHART)
      if (self$has_table) return(MSO_SHAPE_TYPE$TABLE)
      MSO_SHAPE_TYPE$GROUP  # fallback
    },

    # Internal: graphicData URI identifying contents
    .graphic_data_uri = function() {
      nodes <- private$.element$xpath("./a:graphic/a:graphicData/@uri")
      if (length(nodes) == 0) return("")
      xml2::xml_text(nodes[[1]])
    }
  )
)


# ============================================================================
# GroupShapes — shape collection for a group shape
# ============================================================================

#' Shape collection for a group shape
#'
#' @keywords internal
#' @export
GroupShapes <- R6::R6Class(
  "GroupShapes",
  inherit = ParentedElementProxy,

  public = list(
    initialize = function(grpSp, parent) {
      super$initialize(grpSp, parent)
      private$.grpSp <- grpSp
    },

    # Return shape at 1-based index
    get = function(idx) {
      elms <- private$.grpSp$iter_shape_elms()
      if (idx < 1L || idx > length(elms)) {
        stop("shape index out of range", call. = FALSE)
      }
      shape_factory(elms[[idx]], self)
    },

    # All shapes as a list
    to_list = function() {
      lapply(private$.grpSp$iter_shape_elms(), function(e) shape_factory(e, self))
    }
  ),

  private = list(.grpSp = NULL)
)

#' @export
length.GroupShapes <- function(x) {
  length(x$.__enclos_env__$private$.grpSp$iter_shape_elms())
}

#' @export
`[[.GroupShapes` <- function(x, i) x$get(i)
