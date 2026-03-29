# Shape domain objects — BaseShape and concrete subclasses.
#
# Ported from python-pptx/src/pptx/shapes/base.py and subclasses.

# ============================================================================
# shape_factory — dispatch to the correct shape proxy class
# ============================================================================

#' Return appropriate shape proxy for a shape XML element
#'
#' Dispatches on element tag and placeholder context.
#' p:sp with p:ph → SlidePlaceholder / LayoutPlaceholder / MasterPlaceholder;
#' p:sp without p:ph → Shape; p:pic → Picture; p:cxnSp → Connector;
#' p:grpSp → GroupShape; p:graphicFrame → GraphicFrame.
#'
#' @param shape_elm A BaseShapeElement (or subclass) R6 wrapper.
#' @param parent The parent ProvidesPart object (e.g. a Slide).
#' @return An R6 shape proxy.
#' @include enum-shapes.R oxml-shapes.R
#' @keywords internal
#' @export
shape_factory <- function(shape_elm, parent) {
  if (inherits(shape_elm, "CT_Picture"))              return(Picture$new(shape_elm, parent))
  if (inherits(shape_elm, "CT_Connector"))            return(Connector$new(shape_elm, parent))
  if (inherits(shape_elm, "CT_GroupShape"))           return(GroupShape$new(shape_elm, parent))
  if (inherits(shape_elm, "CT_GraphicalObjectFrame")) return(GraphicFrame$new(shape_elm, parent))
  if (inherits(shape_elm, "CT_Shape")) {
    if (isTRUE(shape_elm$has_ph_elm)) {
      if (inherits(parent, "SlideLayout"))  return(LayoutPlaceholder$new(shape_elm, parent))
      if (inherits(parent, "SlideMaster")) return(MasterPlaceholder$new(shape_elm, parent))
      return(SlidePlaceholder$new(shape_elm, parent))
    }
    return(Shape$new(shape_elm, parent))
  }
  BaseShape$new(shape_elm, parent)
}


# ============================================================================
# Layout → master placeholder type mapping
# ============================================================================

# Maps a layout placeholder type string to the corresponding master type.
# Used by LayoutPlaceholder to climb one level higher.
.ph_type_to_master_type <- list(
  body      = "body",
  chart     = "body",
  ctrTitle  = "title",
  dt        = "dt",
  ftr       = "ftr",
  hdr       = "hdr",
  media     = "body",
  obj       = "body",
  dgm       = "body",
  pic       = "body",
  sldImg    = "body",
  sldNum    = "sldNum",
  subTitle  = "body",
  tbl       = "body",
  title     = "title"
)


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

    # ShadowFormat for this shape's shadow effect.
    # shape$shadow$inherit <- FALSE to suppress inherited shadow.
    shadow = function(value) {
      if (!missing(value)) return(invisible(NULL))
      spPr <- private$.element$spPr
      if (is.null(spPr)) stop("shape has no spPr element", call. = FALSE)
      ShadowFormat$new(spPr)
    },

    # ActionSetting for click actions on this shape.
    # No-op setter handles write-back from chaining (shape$click_action$...).
    click_action = function(value) {
      if (!missing(value)) return(invisible(NULL))
      cNvPr <- private$.element$cNvPr
      if (is.null(cNvPr)) stop("shape has no cNvPr element", call. = FALSE)
      ActionSetting$new(cNvPr, self, hover = FALSE)
    },

    # ActionSetting for hover actions on this shape.
    # No-op setter handles write-back from chaining (shape$hover_action$...).
    hover_action = function(value) {
      if (!missing(value)) return(invisible(NULL))
      cNvPr <- private$.element$cNvPr
      if (is.null(cNvPr)) stop("shape has no cNvPr element", call. = FALSE)
      ActionSetting$new(cNvPr, self, hover = TRUE)
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

    # Distance from slide left edge in EMU; inherits from layout placeholder if absent.
    left = function(value) {
      if (!missing(value)) { private$.element$x <- value; return(invisible(value)) }
      if (!is.null(private$.element$xfrm)) return(private$.element$x)
      private$.inherit_dimension("left")
    },

    # Distance from slide top edge in EMU; inherits from layout placeholder if absent.
    top = function(value) {
      if (!missing(value)) { private$.element$y <- value; return(invisible(value)) }
      if (!is.null(private$.element$xfrm)) return(private$.element$y)
      private$.inherit_dimension("top")
    },

    # Shape width in EMU; inherits from layout placeholder if absent.
    width = function(value) {
      if (!missing(value)) { private$.element$cx <- value; return(invisible(value)) }
      if (!is.null(private$.element$xfrm)) return(private$.element$cx)
      private$.inherit_dimension("width")
    },

    # Shape height in EMU; inherits from layout placeholder if absent.
    height = function(value) {
      if (!missing(value)) { private$.element$cy <- value; return(invisible(value)) }
      if (!is.null(private$.element$xfrm)) return(private$.element$cy)
      private$.inherit_dimension("height")
    },

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
  ),

  private = list(
    # Return the layout placeholder corresponding to this slide placeholder,
    # matched by OOXML idx. Returns NULL if not found.
    .base_placeholder = function() {
      if (!self$is_placeholder) return(NULL)
      idx <- private$.element$ph_idx
      slide_part <- tryCatch(private$.parent$part, error = function(e) NULL)
      if (is.null(slide_part)) return(NULL)
      layout <- tryCatch(slide_part$slide_layout, error = function(e) NULL)
      if (is.null(layout)) return(NULL)
      tryCatch(layout$placeholders$get(idx), error = function(e) NULL)
    },

    # Return inherited dimension from base placeholder, or Emu(0L) if none.
    .inherit_dimension = function(dim_name) {
      base <- private$.base_placeholder()
      if (!is.null(base)) return(base[[dim_name]])
      Emu(0L)
    }
  )
)


# ============================================================================
# SlidePlaceholder — placeholder shape on a slide
# ============================================================================

#' Placeholder shape on a slide
#'
#' Inherits dimensions from the corresponding layout placeholder when not
#' explicitly set on the slide element.
#'
#' @keywords internal
#' @export
SlidePlaceholder <- R6::R6Class(
  "SlidePlaceholder",
  inherit = Shape
  # All behaviour inherited from Shape: dimension inheritance from layout,
  # text_frame, text, shape_type == PLACEHOLDER.
)


# ============================================================================
# LayoutPlaceholder — placeholder shape on a slide layout
# ============================================================================

#' Placeholder shape on a slide layout
#'
#' Inherits dimensions from the corresponding master placeholder when not
#' explicitly set on the layout element.
#'
#' @keywords internal
#' @export
LayoutPlaceholder <- R6::R6Class(
  "LayoutPlaceholder",
  inherit = Shape,

  private = list(
    # Return the master placeholder corresponding to this layout placeholder.
    # Maps layout ph_type → master ph_type, then looks up in slide master.
    .base_placeholder = function() {
      if (!self$is_placeholder) return(NULL)
      ph_type <- tryCatch(private$.element$ph_type, error = function(e) NULL)
      if (is.null(ph_type)) return(NULL)
      master_type <- .ph_type_to_master_type[[ph_type]]
      if (is.null(master_type)) return(NULL)
      layout_part <- tryCatch(private$.parent$part, error = function(e) NULL)
      if (is.null(layout_part)) return(NULL)
      master <- tryCatch(layout_part$slide_master, error = function(e) NULL)
      if (is.null(master)) return(NULL)
      tryCatch(master$placeholders$get(master_type), error = function(e) NULL)
    }
  )
)


# ============================================================================
# MasterPlaceholder — placeholder shape on a slide master
# ============================================================================

#' Placeholder shape on a slide master
#'
#' Top of the inheritance chain — dimensions are taken directly from the
#' element; no further inheritance.
#'
#' @keywords internal
#' @export
MasterPlaceholder <- R6::R6Class(
  "MasterPlaceholder",
  inherit = Shape,

  private = list(
    # Master is the top of the chain — no base placeholder.
    .base_placeholder = function() NULL
  )
)


# ============================================================================
# Picture — p:pic
# ============================================================================

# ============================================================================
# PictureCrop — controls the a:srcRect crop on a picture's blipFill
# ============================================================================

#' Picture crop descriptor
#'
#' Provides read/write access to the per-edge crop fractions (0.0–1.0) of a
#' picture. A value of 0.1 means 10% of the edge is cropped.
#' Access via `picture$crop`.
#'
#' @keywords internal
#' @export
PictureCrop <- R6::R6Class(
  "PictureCrop",

  public = list(
    initialize = function(blipFill) private$.blipFill <- blipFill
  ),

  active = list(
    # Fraction cropped from the left edge (0.0-1.0, default 0.0)
    left = function(value) {
      if (!missing(value)) {
        if (value == 0.0 && private$.all_zero_except("l")) {
          private$.blipFill$remove_srcRect()
        } else {
          private$.blipFill$get_or_add_srcRect()$l <- value
        }
        return(invisible(value))
      }
      sr <- private$.blipFill$srcRect
      if (is.null(sr)) 0.0 else sr$l
    },

    # Fraction cropped from the right edge (0.0-1.0, default 0.0)
    right = function(value) {
      if (!missing(value)) {
        if (value == 0.0 && private$.all_zero_except("r")) {
          private$.blipFill$remove_srcRect()
        } else {
          private$.blipFill$get_or_add_srcRect()$r <- value
        }
        return(invisible(value))
      }
      sr <- private$.blipFill$srcRect
      if (is.null(sr)) 0.0 else sr$r
    },

    # Fraction cropped from the top edge (0.0-1.0, default 0.0)
    top = function(value) {
      if (!missing(value)) {
        if (value == 0.0 && private$.all_zero_except("t")) {
          private$.blipFill$remove_srcRect()
        } else {
          private$.blipFill$get_or_add_srcRect()$t <- value
        }
        return(invisible(value))
      }
      sr <- private$.blipFill$srcRect
      if (is.null(sr)) 0.0 else sr$t
    },

    # Fraction cropped from the bottom edge (0.0-1.0, default 0.0)
    bottom = function(value) {
      if (!missing(value)) {
        if (value == 0.0 && private$.all_zero_except("b")) {
          private$.blipFill$remove_srcRect()
        } else {
          private$.blipFill$get_or_add_srcRect()$b <- value
        }
        return(invisible(value))
      }
      sr <- private$.blipFill$srcRect
      if (is.null(sr)) 0.0 else sr$b
    }
  ),

  private = list(
    .blipFill = NULL,

    # Return TRUE if all edges EXCEPT `skip` are zero (used to decide cleanup)
    .all_zero_except = function(skip) {
      sr <- private$.blipFill$srcRect
      if (is.null(sr)) return(TRUE)
      edges <- c(l = sr$l, r = sr$r, t = sr$t, b = sr$b)
      all(edges[setdiff(names(edges), skip)] == 0.0)
    }
  )
)


#' Shape proxy for p:pic elements
#'
#' @keywords internal
#' @export
Picture <- R6::R6Class(
  "Picture",
  inherit = BaseShape,

  active = list(
    shape_type = function() MSO_SHAPE_TYPE$PICTURE,

    # PictureCrop controlling per-edge crop fractions.
    # No-op setter enables R assignment-chain write-back.
    crop = function(value) {
      if (!missing(value)) return(invisible(NULL))
      blipFill <- private$.element$blipFill
      if (is.null(blipFill)) return(NULL)
      PictureCrop$new(blipFill)
    }
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
    shape_type = function() MSO_SHAPE_TYPE$LINE
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

    # rId of the relationship to the chart part, or NULL if none.
    chart_rId = function() {
      node   <- private$.element$get_node()
      c_ns   <- "http://schemas.openxmlformats.org/drawingml/2006/chart"
      r_ns   <- .nsmap[["r"]]
      c_chart <- xml2::xml_find_first(
        node, ".//c:chart",
        ns = c(c = c_ns, r = r_ns)
      )
      if (inherits(c_chart, "xml_missing")) return(NULL)
      val <- xml2::xml_attr(c_chart, "r:id", ns = c(r = r_ns))
      if (is.na(val)) NULL else val
    },

    # Table object for table-containing graphic frames.
    # No-op setter handles write-back from chaining.
    table = function(value) {
      if (!missing(value)) return(invisible(NULL))
      tbl_elm <- private$.element$tbl
      if (is.null(tbl_elm)) stop("this shape does not contain a table", call. = FALSE)
      Table$new(tbl_elm, self)
    },

    # Chart object for chart-containing graphic frames.
    chart = function(value) {
      if (!missing(value)) return(invisible(NULL))
      rId <- self$chart_rId
      if (is.null(rId)) stop("this shape does not contain a chart", call. = FALSE)
      Chart$new(self$part$related_part(rId)$element, self$part$related_part(rId))
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
