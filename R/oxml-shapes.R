# Custom XML element classes for shape-related elements.
#
# Ported from python-pptx/src/pptx/oxml/shapes/shared.py and
# python-pptx/src/pptx/oxml/shapes/groupshape.py.

# ============================================================================
# CT_NonVisualDrawingProps — <p:cNvPr>
# ============================================================================

#' @include oxml-xmlchemy.R oxml-simpletypes.R oxml-init.R
#' @noRd
CT_NonVisualDrawingProps <- define_oxml_element(
  classname = "CT_NonVisualDrawingProps",
  tag = "p:cNvPr",
  attributes = list(
    id   = required_attribute("id",   ST_DrawingElementId),
    name = required_attribute("name", XsdString)
  ),
  active = list(
    # a:hlinkClick child element (click action), or NULL.
    hlinkClick = function(value) {
      if (!missing(value)) return(invisible(NULL))
      nd <- xml2::xml_find_first(self$get_node(), "a:hlinkClick",
                                 ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    # a:hlinkHover child element (hover action), or NULL.
    hlinkHover = function(value) {
      if (!missing(value)) return(invisible(NULL))
      nd <- xml2::xml_find_first(self$get_node(), "a:hlinkHover",
                                 ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  ),
  methods = list(
    get_or_add_hlinkClick = function() {
      nd <- xml2::xml_find_first(self$get_node(), "a:hlinkClick",
                                 ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) {
        nd <- xml2::xml_add_child(self$get_node(), "a:hlinkClick",
                                  xmlns = .nsmap[["a"]])
      }
      wrap_element(nd)
    },
    get_or_add_hlinkHover = function() {
      nd <- xml2::xml_find_first(self$get_node(), "a:hlinkHover",
                                 ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) {
        nd <- xml2::xml_add_child(self$get_node(), "a:hlinkHover",
                                  xmlns = .nsmap[["a"]])
      }
      wrap_element(nd)
    },
    .remove_hlinkClick = function() {
      nd <- xml2::xml_find_first(self$get_node(), "a:hlinkClick",
                                 ns = c(a = .nsmap[["a"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .remove_hlinkHover = function() {
      nd <- xml2::xml_find_first(self$get_node(), "a:hlinkHover",
                                 ns = c(a = .nsmap[["a"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    }
  )
)


# ============================================================================
# CT_ApplicationNonVisualDrawingProps — <p:nvPr>
# ============================================================================

#' @noRd
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

#' @noRd
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

#' @noRd
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

#' @noRd
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

#' @noRd
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

#' @noRd
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
    ),
    ln = zero_or_one(
      "a:ln",
      successors = c("a:effectLst", "a:effectDag", "a:scene3d", "a:sp3d", "a:extLst")
    )
  ),
  active = list(
    # <a:custGeom> child element or NULL
    custGeom = function() {
      nd <- xml2::xml_find_first(self$get_node(), "a:custGeom",
                                 ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    # <a:prstGeom> child element or NULL (read-only)
    prstGeom = function() {
      nd <- xml2::xml_find_first(self$get_node(), "a:prstGeom",
                                 ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_GroupShapeProperties — <p:grpSpPr>
# ============================================================================

#' @noRd
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
#' @noRd
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
    },

    # Return the a:ln element, creating it if absent
    get_or_add_ln = function() {
      spPr <- self$spPr
      if (is.null(spPr)) stop("no spPr element on shape", call. = FALSE)
      spPr$get_or_add_ln()
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

    # The a:ln grandchild element, or NULL
    ln = function() {
      spPr <- self$spPr
      if (is.null(spPr)) return(NULL)
      spPr$ln
    },

    # Shape position and size (EMU, read/write)
    x = function(value) {
      if (!missing(value)) {
        xfrm_elm <- self$get_or_add_xfrm(); xfrm_elm$x <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(Emu(0L)); r <- xfrm$x
      if (is.null(r)) Emu(0L) else r
    },
    y = function(value) {
      if (!missing(value)) {
        xfrm_elm <- self$get_or_add_xfrm(); xfrm_elm$y <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(Emu(0L)); r <- xfrm$y
      if (is.null(r)) Emu(0L) else r
    },
    cx = function(value) {
      if (!missing(value)) {
        xfrm_elm <- self$get_or_add_xfrm(); xfrm_elm$cx <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(Emu(0L)); r <- xfrm$cx
      if (is.null(r)) Emu(0L) else r
    },
    cy = function(value) {
      if (!missing(value)) {
        xfrm_elm <- self$get_or_add_xfrm(); xfrm_elm$cy <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(Emu(0L)); r <- xfrm$cy
      if (is.null(r)) Emu(0L) else r
    },

    # Rotation in clockwise degrees (read/write)
    rot = function(value) {
      if (!missing(value)) {
        xfrm_elm <- self$get_or_add_xfrm(); xfrm_elm$rot <- value; return(invisible(value))
      }
      xfrm <- self$xfrm; if (is.null(xfrm)) return(0.0)
      r <- xfrm$rot; if (is.null(r)) 0.0 else r
    },

    # CT_NonVisualDrawingProps wrapper for this shape's first-child cNvPr, or NULL.
    cNvPr = function() {
      nodes <- self$xpath("./*[1]/p:cNvPr")
      if (length(nodes) == 0) return(NULL)
      wrap_element(nodes[[1]])
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
# CT_Shape factory functions — standalone XML constructors
# ============================================================================

#' Create a new textbox <p:sp> element
#' @noRd
CT_Shape_new_textbox_sp <- function(id, name, x, y, cx, cy) {
  xmlns_a <- .nsmap[["a"]]
  xmlns_p <- .nsmap[["p"]]
  xml_str <- sprintf(
    paste0(
      '<p:sp xmlns:p="%s" xmlns:a="%s">\n',
      '  <p:nvSpPr>\n',
      '    <p:cNvPr id="%d" name="%s"/>\n',
      '    <p:cNvSpPr txBox="1"/>\n',
      '    <p:nvPr/>\n',
      '  </p:nvSpPr>\n',
      '  <p:spPr>\n',
      '    <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>\n',
      '    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>\n',
      '    <a:noFill/>\n',
      '  </p:spPr>\n',
      '  <p:txBody>\n',
      '    <a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr>\n',
      '    <a:lstStyle/>\n',
      '    <a:p/>\n',
      '  </p:txBody>\n',
      '</p:sp>'
    ),
    xmlns_p, xmlns_a,
    as.integer(id), as.character(name),
    as.integer(x), as.integer(y), as.integer(cx), as.integer(cy)
  )
  wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
}

#' Create a new autoshape <p:sp> element
#' @noRd
CT_Shape_new_autoshape_sp <- function(id, name, prst, x, y, cx, cy) {
  xmlns_a <- .nsmap[["a"]]
  xmlns_p <- .nsmap[["p"]]
  xml_str <- sprintf(
    paste0(
      '<p:sp xmlns:p="%s" xmlns:a="%s">\n',
      '  <p:nvSpPr>\n',
      '    <p:cNvPr id="%d" name="%s"/>\n',
      '    <p:cNvSpPr/>\n',
      '    <p:nvPr/>\n',
      '  </p:nvSpPr>\n',
      '  <p:spPr>\n',
      '    <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>\n',
      '    <a:prstGeom prst="%s"><a:avLst/></a:prstGeom>\n',
      '  </p:spPr>\n',
      '  <p:style>\n',
      '    <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>\n',
      '    <a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>\n',
      '    <a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>\n',
      '    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>\n',
      '  </p:style>\n',
      '  <p:txBody>\n',
      '    <a:bodyPr rtlCol="0" anchor="ctr"/>\n',
      '    <a:lstStyle/>\n',
      '    <a:p><a:pPr algn="ctr"/></a:p>\n',
      '  </p:txBody>\n',
      '</p:sp>'
    ),
    xmlns_p, xmlns_a,
    as.integer(id), as.character(name),
    as.integer(x), as.integer(y), as.integer(cx), as.integer(cy),
    as.character(prst)
  )
  wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
}

#' Create a new placeholder <p:sp> element
#' @noRd
CT_Shape_new_placeholder_sp <- function(id, name, ph_type, orient, sz, idx) {
  xmlns_a <- .nsmap[["a"]]
  xmlns_p <- .nsmap[["p"]]
  xml_str <- sprintf(
    paste0(
      '<p:sp xmlns:p="%s" xmlns:a="%s">\n',
      '  <p:nvSpPr>\n',
      '    <p:cNvPr id="%d" name="%s"/>\n',
      '    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>\n',
      '    <p:nvPr/>\n',
      '  </p:nvSpPr>\n',
      '  <p:spPr/>\n',
      '</p:sp>'
    ),
    xmlns_p, xmlns_a,
    as.integer(id), as.character(name)
  )
  sp <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
  # Get nvPr and add the ph element
  nvPr_nodes <- sp$xpath("./*[1]/p:nvPr")
  nvPr <- wrap_element(nvPr_nodes[[1]])
  ph <- nvPr$get_or_add_ph()
  ph$type   <- ph_type
  ph$idx    <- as.integer(idx)
  ph$orient <- orient
  ph$sz     <- sz
  # Add txBody for text-bearing placeholder types
  text_ph_types <- c("title", "ctrTitle", "subTitle", "body", "obj")
  if (!is.null(ph_type) && ph_type %in% text_ph_types) {
    txBody <- .CT_TextBody_new_p_txBody()
    sp$append_child(txBody)
  }
  sp
}


# Escape a string for use in an XML attribute value.
.xml_attr_escape <- function(s) {
  s <- gsub("&",  "&amp;",  s, fixed = TRUE)
  s <- gsub("<",  "&lt;",   s, fixed = TRUE)
  s <- gsub(">",  "&gt;",   s, fixed = TRUE)
  s <- gsub('"',  "&quot;", s, fixed = TRUE)
  s
}


#' Create a new freeform <p:sp> element with custom geometry
#' @noRd
CT_Shape_new_freeform_sp <- function(id, name, x, y, cx, cy) {
  xmlns_a <- .nsmap[["a"]]
  xmlns_p <- .nsmap[["p"]]
  xml_str <- sprintf(
    paste0(
      '<p:sp xmlns:p="%s" xmlns:a="%s">\n',
      '  <p:nvSpPr>\n',
      '    <p:cNvPr id="%d" name="%s"/>\n',
      '    <p:cNvSpPr/>\n',
      '    <p:nvPr/>\n',
      '  </p:nvSpPr>\n',
      '  <p:spPr>\n',
      '    <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>\n',
      '    <a:custGeom>\n',
      '      <a:avLst/>\n',
      '      <a:gdLst/>\n',
      '      <a:ahLst/>\n',
      '      <a:cxnLst/>\n',
      '      <a:rect l="l" t="t" r="r" b="b"/>\n',
      '      <a:pathLst/>\n',
      '    </a:custGeom>\n',
      '  </p:spPr>\n',
      '  <p:style>\n',
      '    <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>\n',
      '    <a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>\n',
      '    <a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>\n',
      '    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>\n',
      '  </p:style>\n',
      '  <p:txBody>\n',
      '    <a:bodyPr rtlCol="0" anchor="ctr"/>\n',
      '    <a:lstStyle/>\n',
      '    <a:p><a:pPr algn="ctr"/></a:p>\n',
      '  </p:txBody>\n',
      '</p:sp>'
    ),
    xmlns_p, xmlns_a,
    as.integer(id), .xml_attr_escape(as.character(name)),
    as.integer(x), as.integer(y), as.integer(cx), as.integer(cy)
  )
  wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
}


#' Create a new <p:cxnSp> connector element
#' @noRd
CT_Connector_new_cxnSp <- function(shape_id, name, prst, x, y, cx, cy,
                                   flipH = FALSE, flipV = FALSE) {
  a_uri <- .nsmap[["a"]]
  p_uri <- .nsmap[["p"]]
  flip  <- paste0(
    if (isTRUE(flipH)) ' flipH="1"' else "",
    if (isTRUE(flipV)) ' flipV="1"' else ""
  )
  xml_str <- sprintf(
    paste0(
      '<p:cxnSp xmlns:p="%s" xmlns:a="%s">\n',
      '  <p:nvCxnSpPr>\n',
      '    <p:cNvPr id="%d" name="%s"/>\n',
      '    <p:cNvCxnSpPr/>\n',
      '    <p:nvPr/>\n',
      '  </p:nvCxnSpPr>\n',
      '  <p:spPr>\n',
      '    <a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>\n',
      '    <a:prstGeom prst="%s"><a:avLst/></a:prstGeom>\n',
      '  </p:spPr>\n',
      '  <p:style>\n',
      '    <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>\n',
      '    <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>\n',
      '    <a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>\n',
      '    <a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>\n',
      '  </p:style>\n',
      '</p:cxnSp>'
    ),
    p_uri, a_uri,
    as.integer(shape_id),
    .xml_attr_escape(as.character(name)),
    flip,
    as.integer(x), as.integer(y), as.integer(cx), as.integer(cy),
    as.character(prst)
  )
  doc <- xml2::read_xml(xml_str)
  wrap_element(xml2::xml_root(doc))
}


#' Create a new <p:pic> picture element
#' @noRd
CT_Picture_new_pic <- function(shape_id, name, desc, rId, x, y, cx, cy) {
  a_uri <- .nsmap[["a"]]
  p_uri <- .nsmap[["p"]]
  r_uri <- .nsmap[["r"]]
  xml_str <- sprintf(
    paste0(
      '<p:pic xmlns:p="%s" xmlns:a="%s" xmlns:r="%s">\n',
      '  <p:nvPicPr>\n',
      '    <p:cNvPr id="%d" name="%s" descr="%s"/>\n',
      '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>\n',
      '    <p:nvPr/>\n',
      '  </p:nvPicPr>\n',
      '  <p:blipFill>\n',
      '    <a:blip r:embed="%s"/>\n',
      '    <a:stretch><a:fillRect/></a:stretch>\n',
      '  </p:blipFill>\n',
      '  <p:spPr>\n',
      '    <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>\n',
      '    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>\n',
      '  </p:spPr>\n',
      '</p:pic>'
    ),
    p_uri, a_uri, r_uri,
    as.integer(shape_id),
    .xml_attr_escape(as.character(name)),
    .xml_attr_escape(as.character(desc)),
    as.character(rId),
    as.integer(x), as.integer(y), as.integer(cx), as.integer(cy)
  )
  doc <- xml2::read_xml(xml_str)
  wrap_element(xml2::xml_root(doc))
}


# ============================================================================
# CT_AdjPoint2D — <a:pt>  (used inside moveTo / lnTo)
# ============================================================================

#' @noRd
CT_AdjPoint2D <- define_oxml_element(
  classname = "CT_AdjPoint2D",
  tag = "a:pt",
  attributes = list(
    x = required_attribute("x", ST_Coordinate),
    y = required_attribute("y", ST_Coordinate)
  )
)


# ============================================================================
# CT_Path2DClose — <a:close>
# ============================================================================

#' @noRd
CT_Path2DClose <- R6::R6Class("CT_Path2DClose", inherit = BaseOxmlElement)


# ============================================================================
# CT_Path2DLineTo — <a:lnTo>
# ============================================================================

#' @noRd
CT_Path2DLineTo <- define_oxml_element(
  classname = "CT_Path2DLineTo",
  tag = "a:lnTo",
  children = list(
    pt = zero_or_one("a:pt", successors = character(0))
  )
)


# ============================================================================
# CT_Path2DMoveTo — <a:moveTo>
# ============================================================================

#' @noRd
CT_Path2DMoveTo <- define_oxml_element(
  classname = "CT_Path2DMoveTo",
  tag = "a:moveTo",
  children = list(
    pt = zero_or_one("a:pt", successors = character(0))
  )
)


# ============================================================================
# CT_Path2D — <a:path>
# ============================================================================

#' @noRd
CT_Path2D <- R6::R6Class(
  "CT_Path2D",
  inherit = BaseOxmlElement,

  public = list(
    # Return newly appended <a:close> child
    add_close = function() {
      nd <- xml2::xml_add_child(self$get_node(), "a:close",
                                xmlns = .nsmap[["a"]])
      wrap_element(nd)
    },

    # Return newly appended <a:lnTo> with embedded <a:pt x y>
    add_lnTo = function(x, y) {
      nd <- xml2::xml_add_child(self$get_node(), "a:lnTo",
                                xmlns = .nsmap[["a"]])
      lnTo <- wrap_element(nd)
      pt_nd <- xml2::xml_add_child(nd, "a:pt",
                                   x = as.character(as.integer(x)),
                                   y = as.character(as.integer(y)),
                                   xmlns = .nsmap[["a"]])
      lnTo
    },

    # Return newly appended <a:moveTo> with embedded <a:pt x y>
    add_moveTo = function(x, y) {
      nd <- xml2::xml_add_child(self$get_node(), "a:moveTo",
                                xmlns = .nsmap[["a"]])
      moveTo <- wrap_element(nd)
      pt_nd <- xml2::xml_add_child(nd, "a:pt",
                                   x = as.character(as.integer(x)),
                                   y = as.character(as.integer(y)),
                                   xmlns = .nsmap[["a"]])
      moveTo
    }
  ),

  active = list(
    # Width of path bounding box (integer EMU or NULL)
    w = function(value) {
      if (!missing(value)) {
        xml2::xml_set_attr(self$get_node(), "w", as.character(as.integer(value)))
        return(invisible(value))
      }
      v <- xml2::xml_attr(self$get_node(), "w")
      if (is.na(v)) NULL else Emu(as.integer(v))
    },

    # Height of path bounding box (integer EMU or NULL)
    h = function(value) {
      if (!missing(value)) {
        xml2::xml_set_attr(self$get_node(), "h", as.character(as.integer(value)))
        return(invisible(value))
      }
      v <- xml2::xml_attr(self$get_node(), "h")
      if (is.na(v)) NULL else Emu(as.integer(v))
    }
  )
)


# ============================================================================
# CT_Path2DList — <a:pathLst>
# ============================================================================

#' @noRd
CT_Path2DList <- R6::R6Class(
  "CT_Path2DList",
  inherit = BaseOxmlElement,

  public = list(
    # Append a new <a:path w h> and return it wrapped
    add_path = function(w, h) {
      nd <- xml2::xml_add_child(self$get_node(), "a:path",
                                w = as.character(as.integer(w)),
                                h = as.character(as.integer(h)),
                                xmlns = .nsmap[["a"]])
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_CustomGeometry2D — <a:custGeom>
# ============================================================================

#' @noRd
CT_CustomGeometry2D <- R6::R6Class(
  "CT_CustomGeometry2D",
  inherit = BaseOxmlElement,

  public = list(
    # Return existing <a:pathLst> or create one
    get_or_add_pathLst = function() {
      nd <- xml2::xml_find_first(self$get_node(), "a:pathLst",
                                 ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) {
        nd <- xml2::xml_add_child(self$get_node(), "a:pathLst",
                                  xmlns = .nsmap[["a"]])
      }
      wrap_element(nd)
    }
  ),

  active = list(
    pathLst = function() {
      nd <- xml2::xml_find_first(self$get_node(), "a:pathLst",
                                 ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
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
#' @noRd
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
    next_shape_id = function() self$max_shape_id() + 1L,

    # Append sp before p:extLst (or at end if none); return the wrapped node
    add_textbox = function(id, name, x, y, cx, cy) {
      sp <- CT_Shape_new_textbox_sp(id, name, x, y, cx, cy)
      self$insert_element_before(sp, "p:extLst")
      sp
    },

    # Append autoshape sp before p:extLst
    add_autoshape = function(id, name, prst, x, y, cx, cy) {
      sp <- CT_Shape_new_autoshape_sp(id, name, prst, x, y, cx, cy)
      self$insert_element_before(sp, "p:extLst")
      sp
    },

    # Append placeholder sp before p:extLst
    add_placeholder = function(id, name, ph_type, orient, sz, idx) {
      sp <- CT_Shape_new_placeholder_sp(id, name, ph_type, orient, sz, idx)
      self$insert_element_before(sp, "p:extLst")
      sp
    },

    # Append a graphicFrame containing a table before p:extLst
    add_table = function(id, name, rows, cols, x, y, cx, cy) {
      gf <- CT_GraphicalObjectFrame_new_table_graphicFrame(
        id, name, rows, cols, x, y, cx, cy
      )
      self$insert_element_before(gf, "p:extLst")
      gf
    },

    # Append a graphicFrame containing a chart reference before p:extLst
    add_chart = function(id, name, rId, x, y, cx, cy) {
      gf <- CT_GraphicalObjectFrame_new_chart_graphicFrame(
        id, name, rId, x, y, cx, cy
      )
      self$insert_element_before(gf, "p:extLst")
      gf
    },

    # Append a p:pic element before p:extLst
    add_pic = function(id, name, desc, rId, x, y, cx, cy) {
      pic <- CT_Picture_new_pic(id, name, desc, rId, x, y, cx, cy)
      self$insert_element_before(pic, "p:extLst")
      pic
    },

    # Append a p:cxnSp connector element before p:extLst
    add_cxnSp = function(id, name, prst, x, y, cx, cy, flipH = FALSE, flipV = FALSE) {
      cxnSp <- CT_Connector_new_cxnSp(id, name, prst, x, y, cx, cy, flipH, flipV)
      self$insert_element_before(cxnSp, "p:extLst")
      cxnSp
    },

    # Append a new freeform p:sp with custom geometry at (x, y, cx, cy)
    add_freeform_sp = function(x, y, cx, cy) {
      shape_id <- self$next_shape_id()
      name     <- sprintf("Freeform %d", shape_id - 1L)
      sp <- CT_Shape_new_freeform_sp(shape_id, name, x, y, cx, cy)
      self$insert_element_before(sp, "p:extLst")
      sp
    },

    # Create a p:grpSp containing shape_elms (list of wrapped elements).
    # Each shape_elm node is removed from its current parent and added to the group.
    add_grpSp = function(id, name, shape_elms) {
      # Compute bounding box across all shapes
      xs  <- sapply(shape_elms, function(e) as.integer(e$x))
      ys  <- sapply(shape_elms, function(e) as.integer(e$y))
      cxs <- sapply(shape_elms, function(e) as.integer(e$cx))
      cys <- sapply(shape_elms, function(e) as.integer(e$cy))
      x   <- min(xs)
      y   <- min(ys)
      cx  <- max(xs + cxs) - x
      cy  <- max(ys + cys) - y
      grpSp <- CT_GroupShape_new_grpSp(id, name, x, y, cx, cy)
      # Move shape nodes into the group
      for (elm in shape_elms) {
        xml2::xml_remove(elm$get_node())
        xml2::xml_add_child(grpSp$get_node(), elm$get_node())
      }
      self$insert_element_before(grpSp, "p:extLst")
      grpSp
    }
  ),

  active = list(
    # Override: group uses p:grpSpPr, not p:spPr
    spPr = function() self$find(qn("p:grpSpPr"))
  )
)


# ============================================================================
# Concrete shape element classes — p:sp, p:pic, p:cxnSp, p:graphicFrame
# ============================================================================

#' CT_Shape XML element
#' @noRd
#' @export
CT_Shape <- define_oxml_element(
  classname = "CT_Shape",
  tag = "p:sp",
  children = list(
    txBody = zero_or_one("p:txBody", successors = character(0))
  ),
  inherit = BaseShapeElement
)

# Override auto-generated _new_txBody to produce a properly structured <p:txBody>.
CT_Shape$set("public", "_new_txBody", function() .CT_TextBody_new_p_txBody(), overwrite = TRUE)

# Add freeform-related methods to CT_Shape
CT_Shape$set("public", "add_path", function(w, h) {
  custGeom <- self$spPr$custGeom
  if (is.null(custGeom)) stop("shape must be freeform (no custGeom)", call. = FALSE)
  pathLst <- custGeom$get_or_add_pathLst()
  pathLst$add_path(w = w, h = h)
}, overwrite = TRUE)

CT_Shape$set("active", "has_custom_geometry", function() {
  spPr <- self$spPr
  if (is.null(spPr)) return(FALSE)
  !is.null(spPr$custGeom)
}, overwrite = TRUE)

# ============================================================================
# CT_SrcRect — <a:srcRect> — per-edge crop percentages on a blipFill
# ============================================================================

#' @noRd
CT_SrcRect <- R6::R6Class(
  "CT_SrcRect",
  inherit = BaseOxmlElement,
  active = list(
    l = function(value) {
      if (!missing(value)) {
        v <- as.integer(round(as.numeric(value) * 100000.0))
        if (v == 0L) self$remove_attr("l") else self$set_attr("l", as.character(v))
        return(invisible(value))
      }
      v <- self$get_attr("l"); if (is.null(v)) 0.0 else as.integer(v) / 100000.0
    },
    r = function(value) {
      if (!missing(value)) {
        v <- as.integer(round(as.numeric(value) * 100000.0))
        if (v == 0L) self$remove_attr("r") else self$set_attr("r", as.character(v))
        return(invisible(value))
      }
      v <- self$get_attr("r"); if (is.null(v)) 0.0 else as.integer(v) / 100000.0
    },
    t = function(value) {
      if (!missing(value)) {
        v <- as.integer(round(as.numeric(value) * 100000.0))
        if (v == 0L) self$remove_attr("t") else self$set_attr("t", as.character(v))
        return(invisible(value))
      }
      v <- self$get_attr("t"); if (is.null(v)) 0.0 else as.integer(v) / 100000.0
    },
    b = function(value) {
      if (!missing(value)) {
        v <- as.integer(round(as.numeric(value) * 100000.0))
        if (v == 0L) self$remove_attr("b") else self$set_attr("b", as.character(v))
        return(invisible(value))
      }
      v <- self$get_attr("b"); if (is.null(v)) 0.0 else as.integer(v) / 100000.0
    }
  )
)


# ============================================================================
# CT_BlipFill — <p:blipFill>
# ============================================================================

#' @noRd
CT_BlipFill <- R6::R6Class(
  "CT_BlipFill",
  inherit = BaseOxmlElement,

  public = list(
    # Return or create <a:srcRect>
    get_or_add_srcRect = function() {
      existing <- self$srcRect
      if (!is.null(existing)) return(existing)
      # Insert before a:stretch (or append)
      stretch_nd <- xml2::xml_find_first(
        self$get_node(), "a:stretch", ns = c(a = .nsmap[["a"]])
      )
      new_nd <- xml2::read_xml(sprintf('<a:srcRect xmlns:a="%s"/>', .nsmap[["a"]]))
      if (!inherits(stretch_nd, "xml_missing")) {
        xml2::xml_add_sibling(stretch_nd, xml2::xml_root(new_nd), .where = "before")
      } else {
        xml2::xml_add_child(self$get_node(), xml2::xml_root(new_nd))
      }
      self$srcRect
    },

    # Remove <a:srcRect> if present
    remove_srcRect = function() {
      nd <- xml2::xml_find_first(
        self$get_node(), "a:srcRect", ns = c(a = .nsmap[["a"]])
      )
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
      invisible(NULL)
    }
  ),

  active = list(
    srcRect = function() {
      nd <- xml2::xml_find_first(
        self$get_node(), "a:srcRect", ns = c(a = .nsmap[["a"]])
      )
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_Picture XML element
# ============================================================================

#' CT_Picture XML element
#' @noRd
#' @export
CT_Picture <- R6::R6Class(
  "CT_Picture",
  inherit = BaseShapeElement,
  active = list(
    blipFill = function() {
      nd <- xml2::xml_find_first(
        self$get_node(), "p:blipFill", ns = c(p = .nsmap[["p"]])
      )
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)

#' CT_Connector XML element
#' @noRd
#' @export
CT_Connector <- R6::R6Class(
  "CT_Connector",
  inherit = BaseShapeElement,

  active = list(
    # Return the <p:cNvCxnSpPr> element (always present in a valid cxnSp)
    cNvCxnSpPr = function() {
      nd <- xml2::xml_find_first(
        self$get_node(), "p:nvCxnSpPr/p:cNvCxnSpPr",
        ns = c(p = .nsmap[["p"]])
      )
      if (inherits(nd, "xml_missing")) return(NULL)
      nd
    },

    # The <a:stCxn> child of cNvCxnSpPr, or NULL.
    stCxn = function() {
      pr <- self$cNvCxnSpPr
      if (is.null(pr)) return(NULL)
      nd <- xml2::xml_find_first(pr, "a:stCxn", ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      nd
    },

    # The <a:endCxn> child of cNvCxnSpPr, or NULL.
    endCxn = function() {
      pr <- self$cNvCxnSpPr
      if (is.null(pr)) return(NULL)
      nd <- xml2::xml_find_first(pr, "a:endCxn", ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      nd
    }
  ),

  public = list(
    # Set (or replace) the <a:stCxn id="shapeId" idx="siteIdx"/> element.
    set_stCxn = function(shape_id, site_idx) {
      pr <- self$cNvCxnSpPr
      if (is.null(pr)) stop("cxnSp has no cNvCxnSpPr", call. = FALSE)
      nd <- xml2::xml_find_first(pr, "a:stCxn", ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) {
        xml_str <- sprintf('<a:stCxn xmlns:a="%s" id="%d" idx="%d"/>',
                           .nsmap[["a"]], as.integer(shape_id), as.integer(site_idx))
        xml2::xml_add_child(pr, xml2::xml_root(xml2::read_xml(xml_str)), .where = 0L)
      } else {
        xml2::xml_set_attr(nd, "id",  as.character(shape_id))
        xml2::xml_set_attr(nd, "idx", as.character(site_idx))
      }
      invisible(self)
    },

    # Set (or replace) the <a:endCxn id="shapeId" idx="siteIdx"/> element.
    set_endCxn = function(shape_id, site_idx) {
      pr <- self$cNvCxnSpPr
      if (is.null(pr)) stop("cxnSp has no cNvCxnSpPr", call. = FALSE)
      nd <- xml2::xml_find_first(pr, "a:endCxn", ns = c(a = .nsmap[["a"]]))
      if (inherits(nd, "xml_missing")) {
        xml_str <- sprintf('<a:endCxn xmlns:a="%s" id="%d" idx="%d"/>',
                           .nsmap[["a"]], as.integer(shape_id), as.integer(site_idx))
        xml2::xml_add_child(pr, xml2::xml_root(xml2::read_xml(xml_str)))
      } else {
        xml2::xml_set_attr(nd, "id",  as.character(shape_id))
        xml2::xml_set_attr(nd, "idx", as.character(site_idx))
      }
      invisible(self)
    },

    # Remove the <a:stCxn> element (disconnect begin endpoint).
    remove_stCxn = function() {
      nd <- self$stCxn
      if (!is.null(nd)) xml2::xml_remove(nd)
      invisible(self)
    },

    # Remove the <a:endCxn> element (disconnect end endpoint).
    remove_endCxn = function() {
      nd <- self$endCxn
      if (!is.null(nd)) xml2::xml_remove(nd)
      invisible(self)
    }
  )
)

#' CT_GraphicalObjectFrame XML element
#' @noRd
#' @export
CT_GraphicalObjectFrame <- R6::R6Class("CT_GraphicalObjectFrame", inherit = BaseShapeElement)


# ============================================================================
# Factory — create a graphicFrame containing a chart reference
# ============================================================================

.GRAPHIC_DATA_URI_CHART <- "http://schemas.openxmlformats.org/drawingml/2006/chart"

#' Create a new <p:graphicFrame> referencing a chart part via rId
#' @noRd
CT_GraphicalObjectFrame_new_chart_graphicFrame <- function(id, name, rId, x, y, cx, cy) {
  p <- .nsmap[["p"]]; a <- .nsmap[["a"]]
  c_ns <- "http://schemas.openxmlformats.org/drawingml/2006/chart"
  r_ns <- .nsmap[["r"]]
  xml_str <- sprintf(paste0(
    '<p:graphicFrame xmlns:p="%s" xmlns:a="%s" xmlns:r="%s">',
      '<p:nvGraphicFramePr>',
        '<p:cNvPr id="%d" name="%s"/>',
        '<p:cNvGraphicFramePr/>',
        '<p:nvPr/>',
      '</p:nvGraphicFramePr>',
      '<p:xfrm>',
        '<a:off x="%d" y="%d"/>',
        '<a:ext cx="%d" cy="%d"/>',
      '</p:xfrm>',
      '<a:graphic>',
        '<a:graphicData uri="%s">',
          '<c:chart xmlns:c="%s" r:id="%s"/>',
        '</a:graphicData>',
      '</a:graphic>',
    '</p:graphicFrame>'
  ), p, a, r_ns,
    as.integer(id), as.character(name),
    as.integer(x), as.integer(y), as.integer(cx), as.integer(cy),
    .GRAPHIC_DATA_URI_CHART,
    c_ns, rId)

  wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
}


# ============================================================================
# Factory — create a p:grpSp (group shape) element
# ============================================================================

#' Create a new empty <p:grpSp> with bounding-box transform
#' @noRd
CT_GroupShape_new_grpSp <- function(id, name, x, y, cx, cy) {
  p <- .nsmap[["p"]]; a <- .nsmap[["a"]]
  xml_str <- sprintf(paste0(
    '<p:grpSp xmlns:p="%s" xmlns:a="%s">',
      '<p:nvGrpSpPr>',
        '<p:cNvPr id="%d" name="%s"/>',
        '<p:cNvGrpSpPr/>',
        '<p:nvPr/>',
      '</p:nvGrpSpPr>',
      '<p:grpSpPr>',
        '<a:xfrm>',
          '<a:off x="%d" y="%d"/>',
          '<a:ext cx="%d" cy="%d"/>',
          '<a:chOff x="%d" y="%d"/>',
          '<a:chExt cx="%d" cy="%d"/>',
        '</a:xfrm>',
      '</p:grpSpPr>',
    '</p:grpSp>'
  ), p, a,
     as.integer(id), as.character(name),
     as.integer(x), as.integer(y), as.integer(cx), as.integer(cy),
     as.integer(x), as.integer(y), as.integer(cx), as.integer(cy))
  wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
}


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
  register_element_cls("p:blipFill",      CT_BlipFill)
  register_element_cls("a:srcRect",      CT_SrcRect)
  register_element_cls("p:pic",          CT_Picture)
  register_element_cls("p:cxnSp",        CT_Connector)
  register_element_cls("p:graphicFrame", CT_GraphicalObjectFrame)
  # Freeform path elements
  register_element_cls("a:custGeom",     CT_CustomGeometry2D)
  register_element_cls("a:pathLst",      CT_Path2DList)
  register_element_cls("a:path",         CT_Path2D)
  register_element_cls("a:pt",           CT_AdjPoint2D)
  register_element_cls("a:close",        CT_Path2DClose)
  register_element_cls("a:lnTo",         CT_Path2DLineTo)
  register_element_cls("a:moveTo",       CT_Path2DMoveTo)
}
