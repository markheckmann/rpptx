# Custom XML element classes for shape-related elements.
#
# Ported from python-pptx/src/pptx/oxml/shapes/shared.py and
# python-pptx/src/pptx/oxml/shapes/groupshape.py.

# ============================================================================
# CT_NonVisualDrawingProps — <p:cNvPr>
# ============================================================================

#' @include oxml-xmlchemy.R oxml-simpletypes.R oxml-init.R
#' @keywords internal
CT_NonVisualDrawingProps <- define_oxml_element(
  classname = "CT_NonVisualDrawingProps",
  tag = "p:cNvPr",
  attributes = list(
    id   = required_attribute("id",   ST_DrawingElementId),
    name = required_attribute("name", XsdString)
  )
)


# ============================================================================
# CT_ApplicationNonVisualDrawingProps — <p:nvPr>
# ============================================================================

#' @keywords internal
CT_ApplicationNonVisualDrawingProps <- define_oxml_element(
  classname = "CT_ApplicationNonVisualDrawingProps",
  tag = "p:nvPr",
  children = list(
    ph = zero_or_one(
      "p:ph",
      successors = c("a:audioCd", "a:wavAudioFile", "a:audioFile",
                     "a:videoFile", "a:quickTimeFile", "p:custDataLst", "p:extLst")
    )
  )
)


# ============================================================================
# CT_Placeholder — <p:ph>
# ============================================================================

#' @keywords internal
CT_Placeholder <- define_oxml_element(
  classname = "CT_Placeholder",
  tag = "p:ph",
  attributes = list(
    type   = optional_attribute("type",   XsdString,      default = "obj"),
    orient = optional_attribute("orient", XsdString,      default = "horz"),
    sz     = optional_attribute("sz",     XsdString,      default = "full"),
    idx    = optional_attribute("idx",    XsdUnsignedInt, default = 0L)
  )
)


# ============================================================================
# CT_Point2D — <a:off>
# ============================================================================

#' @keywords internal
CT_Point2D <- define_oxml_element(
  classname = "CT_Point2D",
  tag = "a:off",
  attributes = list(
    x = required_attribute("x", ST_Coordinate),
    y = required_attribute("y", ST_Coordinate)
  )
)


# ============================================================================
# CT_PositiveSize2D — <a:ext>
# ============================================================================

#' @keywords internal
CT_PositiveSize2D <- define_oxml_element(
  classname = "CT_PositiveSize2D",
  tag = "a:ext",
  attributes = list(
    cx = required_attribute("cx", ST_PositiveCoordinate),
    cy = required_attribute("cy", ST_PositiveCoordinate)
  )
)


# ============================================================================
# CT_Transform2D — <a:xfrm>
# ============================================================================

#' @keywords internal
CT_Transform2D <- define_oxml_element(
  classname = "CT_Transform2D",
  tag = "a:xfrm",
  children = list(
    off   = zero_or_one("a:off",   successors = c("a:ext", "a:chOff", "a:chExt")),
    ext   = zero_or_one("a:ext",   successors = c("a:chOff", "a:chExt")),
    chOff = zero_or_one("a:chOff", successors = c("a:chExt")),
    chExt = zero_or_one("a:chExt")
  ),
  attributes = list(
    rot   = optional_attribute("rot",   ST_Angle,   default = 0.0),
    flipH = optional_attribute("flipH", XsdBoolean, default = FALSE),
    flipV = optional_attribute("flipV", XsdBoolean, default = FALSE)
  ),
  active = list(
    # Delegate x/y to off child, cx/cy to ext child
    x = function(value) {
      if (!missing(value)) {
        off <- self$get_or_add_off(); off$x <- value; return(invisible(value))
      }
      off <- self$off; if (is.null(off)) NULL else off$x
    },
    y = function(value) {
      if (!missing(value)) {
        off <- self$get_or_add_off(); off$y <- value; return(invisible(value))
      }
      off <- self$off; if (is.null(off)) NULL else off$y
    },
    cx = function(value) {
      if (!missing(value)) {
        ext <- self$get_or_add_ext(); ext$cx <- value; return(invisible(value))
      }
      ext <- self$ext; if (is.null(ext)) NULL else ext$cx
    },
    cy = function(value) {
      if (!missing(value)) {
        ext <- self$get_or_add_ext(); ext$cy <- value; return(invisible(value))
      }
      ext <- self$ext; if (is.null(ext)) NULL else ext$cy
    }
  )
)



# ============================================================================
# CT_ShapeProperties — <p:spPr>
# ============================================================================

#' @keywords internal
CT_ShapeProperties <- define_oxml_element(
  classname = "CT_ShapeProperties",
  tag = "p:spPr",
  children = list(
    xfrm = zero_or_one(
      "a:xfrm",
      successors = c("a:custGeom", "a:prstGeom",
                     "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill",
                     "a:pattFill", "a:grpFill", "a:ln",
                     "a:effectLst", "a:effectDag", "a:scene3d", "a:sp3d", "a:extLst")
    )
  )
)


# ============================================================================
# CT_GroupShapeProperties — <p:grpSpPr>
# ============================================================================

#' @keywords internal
CT_GroupShapeProperties <- define_oxml_element(
  classname = "CT_GroupShapeProperties",
  tag = "p:grpSpPr",
  children = list(
    xfrm = zero_or_one(
      "a:xfrm",
      successors = c("a:noFill", "a:solidFill", "a:gradFill",
                     "a:blipFill", "a:pattFill", "a:grpFill",
                     "a:effectLst", "a:effectDag", "a:scene3d", "a:extLst")
    )
  )
)


# ============================================================================
# BaseShapeElement — common base for all shape XML elements
# ============================================================================

#' Base class for shape element classes (p:sp, p:pic, p:cxnSp, p:grpSp, etc.)
#'
#' Provides position, size, rotation, shape ID/name, and placeholder detection
#' via active bindings that delegate through the element's shape properties.
#'
#' @keywords internal
#' @export
BaseShapeElement <- R6::R6Class(
  "BaseShapeElement",
  inherit = BaseOxmlElement,

  public = list(
    # Return the a:xfrm element, creating it (and spPr) if absent
    get_or_add_xfrm = function() {
      spPr <- self$spPr
      if (is.null(spPr)) stop("no spPr element on shape", call. = FALSE)
      spPr$get_or_add_xfrm()
    }
  ),

  active = list(
    # The shape-properties element (p:spPr for most shapes; override for groups)
    spPr = function() self$find(qn("p:spPr")),

    # The a:xfrm grandchild element, or NULL
    xfrm = function() {
      spPr <- self$spPr
      if (is.null(spPr)) return(NULL)
      spPr$xfrm
    },

    # Shape position and size (EMU, read/write)
    x = function(value) {
      if (!missing(value)) {
        self$get_or_add_xfrm()$x <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(Emu(0L)); r <- xfrm$x
      if (is.null(r)) Emu(0L) else r
    },
    y = function(value) {
      if (!missing(value)) {
        self$get_or_add_xfrm()$y <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(Emu(0L)); r <- xfrm$y
      if (is.null(r)) Emu(0L) else r
    },
    cx = function(value) {
      if (!missing(value)) {
        self$get_or_add_xfrm()$cx <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(Emu(0L)); r <- xfrm$cx
      if (is.null(r)) Emu(0L) else r
    },
    cy = function(value) {
      if (!missing(value)) {
        self$get_or_add_xfrm()$cy <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(Emu(0L)); r <- xfrm$cy
      if (is.null(r)) Emu(0L) else r
    },

    # Rotation in clockwise degrees (read/write)
    rot = function(value) {
      if (!missing(value)) {
        self$get_or_add_xfrm()$rot <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(0.0)
      r <- xfrm$rot; if (is.null(r)) 0.0 else r
    },

    # Integer shape ID (from p:cNvPr/@id of first child)
    shape_id = function() {
      nodes <- self$xpath("./*[1]/p:cNvPr")
      if (length(nodes) == 0) return(NULL)
      id_str <- xml2::xml_attr(nodes[[1]], "id")
      if (is.na(id_str)) return(NULL)
      as.integer(id_str)
    },

    # Shape name string (from p:cNvPr/@name of first child)
    shape_name = function(value) {
      nodes <- self$xpath("./*[1]/p:cNvPr")
      if (length(nodes) == 0) return(NULL)
      if (!missing(value)) {
        xml2::xml_set_attr(nodes[[1]], "name", as.character(value))
        return(invisible(value))
      }
      xml2::xml_attr(nodes[[1]], "name")
    },

    # TRUE if this element has a <p:ph> descendant (is a placeholder)
    has_ph_elm = function() !is.null(self$ph),

    # The <p:ph> element, or NULL (via nvXxPr > nvPr > ph)
    ph = function() {
      nodes <- self$xpath("./*[1]/p:nvPr/p:ph")
      if (length(nodes) == 0) return(NULL)
      wrap_element(nodes[[1]])
    },

    # Placeholder idx attribute (raises error if not a placeholder)
    ph_idx = function() {
      ph <- self$ph
      if (is.null(ph)) stop("not a placeholder shape", call. = FALSE)
      ph$idx
    },

    # Placeholder type string (raises error if not a placeholder)
    ph_type = function() {
      ph <- self$ph
      if (is.null(ph)) stop("not a placeholder shape", call. = FALSE)
      ph$type
    },

    # Placeholder orientation (raises error if not a placeholder)
    ph_orient = function() {
      ph <- self$ph
      if (is.null(ph)) stop("not a placeholder shape", call. = FALSE)
      ph$orient
    },

    # Placeholder size (raises error if not a placeholder)
    ph_sz = function() {
      ph <- self$ph
      if (is.null(ph)) stop("not a placeholder shape", call. = FALSE)
      ph$sz
    }
  )
)


# ============================================================================
# CT_GroupShape — <p:spTree> and <p:grpSp>
# ============================================================================

# Shape tags used for iteration
.shape_tags <- c(
  qn("p:sp"), qn("p:grpSp"), qn("p:graphicFrame"),
  qn("p:cxnSp"), qn("p:pic"), qn("p:contentPart")
)

#' Custom element class for p:spTree and p:grpSp elements
#'
#' @keywords internal
#' @export
CT_GroupShape <- R6::R6Class(
  "CT_GroupShape",
  inherit = BaseShapeElement,

  public = list(
    # Return list of child shape elements (wrapped)
    iter_shape_elms = function() {
      children <- xml2::xml_children(private$.node)
      shapes <- list()
      for (child in children) {
        clark <- .get_clark_name(child)
        if (clark %in% .shape_tags) {
          shapes <- c(shapes, list(wrap_element(child)))
        }
      }
      shapes
    },

    # Return list of placeholder shape children
    iter_ph_elms = function() {
      Filter(function(e) e$has_ph_elm, self$iter_shape_elms())
    },

    # Maximum @id value anywhere in the document (integer)
    max_shape_id = function() {
      id_nodes <- xml2::xml_find_all(
        xml2::xml_root(private$.node), "//@id"
      )
      ids <- suppressWarnings(as.integer(xml2::xml_text(id_nodes)))
      ids <- ids[!is.na(ids)]
      if (length(ids) == 0L) return(0L)
      max(ids)
    },

    # Next unique shape id (max + 1)
    next_shape_id = function() self$max_shape_id() + 1L
  ),

  active = list(
    # Override: group uses p:grpSpPr, not p:spPr
    spPr = function() self$find(qn("p:grpSpPr"))
  )
)


# ============================================================================
# Concrete shape element classes — p:sp, p:pic, p:cxnSp, p:graphicFrame
# ============================================================================

#' @keywords internal
#' @export
CT_Shape <- R6::R6Class("CT_Shape", inherit = BaseShapeElement)

#' @keywords internal
#' @export
CT_Picture <- R6::R6Class("CT_Picture", inherit = BaseShapeElement)

#' @keywords internal
#' @export
CT_Connector <- R6::R6Class("CT_Connector", inherit = BaseShapeElement)

#' @keywords internal
#' @export
CT_GraphicalObjectFrame <- R6::R6Class("CT_GraphicalObjectFrame", inherit = BaseShapeElement)


# ============================================================================
# Register element classes
# ============================================================================

.onLoad_oxml_shapes <- function() {
  register_element_cls("p:cNvPr",        CT_NonVisualDrawingProps)
  register_element_cls("p:nvPr",         CT_ApplicationNonVisualDrawingProps)
  register_element_cls("p:ph",           CT_Placeholder)
  register_element_cls("a:off",          CT_Point2D)
  register_element_cls("a:ext",          CT_PositiveSize2D)
  register_element_cls("a:chOff",        CT_Point2D)
  register_element_cls("a:chExt",        CT_PositiveSize2D)
  register_element_cls("a:xfrm",         CT_Transform2D)
  register_element_cls("p:spPr",         CT_ShapeProperties)
  register_element_cls("p:grpSpPr",      CT_GroupShapeProperties)
  register_element_cls("p:spTree",       CT_GroupShape)
  register_element_cls("p:grpSp",        CT_GroupShape)
  register_element_cls("p:sp",           CT_Shape)
  register_element_cls("p:pic",          CT_Picture)
  register_element_cls("p:cxnSp",        CT_Connector)
  register_element_cls("p:graphicFrame", CT_GraphicalObjectFrame)
}
