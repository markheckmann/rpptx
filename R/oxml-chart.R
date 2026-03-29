# Chart oxml element wrappers.
#
# Ported from python-pptx/src/pptx/oxml/chart/*.py


# ============================================================================
# Helpers shared across chart elements
# ============================================================================

.chart_ns_decls <- function(...) {
  nms <- c("c", "a", "r", ...)
  unique_nms <- unique(nms)
  paste(vapply(unique_nms, function(nm) {
    sprintf('xmlns:%s="%s"', nm, .nsmap[[nm]])
  }, character(1)), collapse = " ")
}


# ============================================================================
# CT_Boolean — <c:*> elements with a boolean val attribute
# ============================================================================

#' Boolean chart element (val defaults to TRUE)
#' @noRd
#' @export
CT_Boolean <- R6::R6Class(
  "CT_Boolean",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", if (isTRUE(value)) "1" else "0")
        return(invisible(value))
      }
      v <- self$get_attr("val")
      if (is.null(v) || is.na(v)) return(TRUE)   # default TRUE per schema
      !(v %in% c("0", "false", "FALSE"))
    }
  )
)


# ============================================================================
# CT_Double — <c:*> elements with a required numeric val attribute
# ============================================================================

#' Numeric val chart element
#' @noRd
#' @export
CT_Double <- R6::R6Class(
  "CT_Double",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(as.numeric(value)))
        return(invisible(value))
      }
      as.numeric(self$get_attr("val"))
    }
  )
)


# ============================================================================
# CT_UnsignedInt — <c:idx>, <c:order> etc.
# ============================================================================

#' Unsigned integer val chart element
#' @noRd
#' @export
CT_UnsignedInt <- R6::R6Class(
  "CT_UnsignedInt",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(as.integer(value)))
        return(invisible(value))
      }
      as.integer(self$get_attr("val"))
    }
  )
)


# ============================================================================
# CT_AxisUnit — <c:majorUnit>, <c:minorUnit>
# ============================================================================

#' Axis unit element (numeric val)
#' @noRd
#' @export
CT_AxisUnit <- R6::R6Class(
  "CT_AxisUnit",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(as.numeric(value)))
        return(invisible(value))
      }
      as.numeric(self$get_attr("val"))
    }
  )
)


# ============================================================================
# CT_NumFmt — <c:numFmt>
# ============================================================================

#' Number format element
#' @noRd
#' @export
CT_NumFmt <- R6::R6Class(
  "CT_NumFmt",
  inherit = BaseOxmlElement,
  active = list(
    formatCode = function(value) {
      if (!missing(value)) {
        self$set_attr("formatCode", as.character(value))
        return(invisible(value))
      }
      self$get_attr("formatCode")
    },
    sourceLinked = function(value) {
      if (!missing(value)) {
        self$set_attr("sourceLinked", if (isTRUE(value)) "1" else "0")
        return(invisible(value))
      }
      v <- self$get_attr("sourceLinked")
      if (is.null(v) || is.na(v)) return(NULL)
      !(v %in% c("0", "false", "FALSE"))
    }
  )
)


# ============================================================================
# CT_Scaling — <c:scaling>
# ============================================================================

#' Axis scaling element
#' @noRd
#' @export
CT_Scaling <- R6::R6Class(
  "CT_Scaling",
  inherit = BaseOxmlElement,

  active = list(
    orientation = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:orientation",
                                  ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    max = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:max",
                                  ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    min = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:min",
                                  ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    maximum = function(value) {
      if (!missing(value)) {
        self$.remove_max()
        if (!is.null(value)) self$.add_max(val = value)
        return(invisible(value))
      }
      mx <- self$max
      if (is.null(mx)) return(NULL)
      mx$val
    },

    minimum = function(value) {
      if (!missing(value)) {
        self$.remove_min()
        if (!is.null(value)) self$.add_min(val = value)
        return(invisible(value))
      }
      mn <- self$min
      if (is.null(mn)) return(NULL)
      mn$val
    }
  ),

  public = list(
    .remove_max = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:max", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .add_max = function(val) {
      node <- self$get_node()
      mx <- xml2::xml_add_child(node, "c:max", xmlns = .nsmap[["c"]])
      xml2::xml_set_attr(mx, "val", as.character(as.numeric(val)))
    },
    .remove_min = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:min", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .add_min = function(val) {
      node <- self$get_node()
      mn <- xml2::xml_add_child(node, "c:min", xmlns = .nsmap[["c"]])
      xml2::xml_set_attr(mn, "val", as.character(as.numeric(val)))
    },
    get_or_add_orientation = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:orientation", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) {
        nd <- xml2::xml_add_child(self$get_node(), "c:orientation", xmlns = .nsmap[["c"]])
      }
      wrap_element(nd)
    },
    .remove_orientation = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:orientation", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    }
  )
)


# ============================================================================
# CT_Orientation — <c:orientation val="minMax|maxMin">
# ============================================================================

#' Axis orientation element
#' @noRd
#' @export
CT_Orientation <- R6::R6Class(
  "CT_Orientation",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(value))
        return(invisible(value))
      }
      v <- self$get_attr("val")
      if (is.null(v) || is.na(v)) return("minMax")
      v
    }
  )
)


# ============================================================================
# CT_LblOffset — <c:lblOffset val="...">
# ============================================================================

#' Label offset element (default 100)
#' @noRd
#' @export
CT_LblOffset <- R6::R6Class(
  "CT_LblOffset",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(as.integer(value)))
        return(invisible(value))
      }
      v <- self$get_attr("val")
      if (is.null(v) || is.na(v)) return(100L)
      as.integer(sub("%", "", v))
    }
  )
)


# ============================================================================
# CT_TickLblPos — <c:tickLblPos val="...">
# ============================================================================

#' Tick label position element
#' @noRd
#' @export
CT_TickLblPos <- R6::R6Class(
  "CT_TickLblPos",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(value))
        return(invisible(value))
      }
      self$get_attr("val")
    }
  )
)


# ============================================================================
# CT_TickMark — <c:majorTickMark>, <c:minorTickMark>
# ============================================================================

#' Tick mark element (cross/in/none/out)
#' @noRd
#' @export
CT_TickMark <- R6::R6Class(
  "CT_TickMark",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(value))
        return(invisible(value))
      }
      v <- self$get_attr("val")
      if (is.null(v) || is.na(v)) return(XL_TICK_MARK$CROSS)
      v
    }
  )
)


# ============================================================================
# CT_Crosses — <c:crosses val="autoZero|max|min">
# ============================================================================

#' Axis crossing element
#' @noRd
#' @export
CT_Crosses <- R6::R6Class(
  "CT_Crosses",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(value))
        return(invisible(value))
      }
      self$get_attr("val")
    }
  )
)


# ============================================================================
# CT_CrossesAt — <c:crossesAt val="float">
# ============================================================================

#' Axis crosses-at element (numeric)
#' @noRd
#' @export
CT_CrossesAt <- R6::R6Class(
  "CT_CrossesAt",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(as.numeric(value)))
        return(invisible(value))
      }
      as.numeric(self$get_attr("val"))
    }
  )
)


# ============================================================================
# BaseAxisElement — shared base for catAx, valAx, dateAx
# ============================================================================

#' Base class for chart axis XML elements
#' @noRd
#' @export
BaseAxisElement <- R6::R6Class(
  "BaseAxisElement",
  inherit = BaseOxmlElement,

  public = list(
    # Get or create txPr child
    get_or_add_txPr = function() {
      .get_or_add_child(self$get_node(), "c:txPr", .nsmap[["c"]])
    },
    get_or_add_title = function() {
      nd <- .get_or_add_child(self$get_node(), "c:title", .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_title = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:title", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    get_or_add_majorGridlines = function() {
      nd <- .get_or_add_child(self$get_node(), "c:majorGridlines", .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_majorGridlines = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:majorGridlines", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    get_or_add_minorGridlines = function() {
      nd <- .get_or_add_child(self$get_node(), "c:minorGridlines", .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_minorGridlines = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:minorGridlines", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    get_or_add_tickLblPos = function() {
      nd <- .get_or_add_child(self$get_node(), "c:tickLblPos", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_numFmt = function() {
      nd <- .get_or_add_child(self$get_node(), "c:numFmt", .nsmap[["c"]])
      wrap_element(nd)
    },
    .add_majorTickMark = function(val = NULL) {
      nd <- xml2::xml_add_child(self$get_node(), "c:majorTickMark", xmlns = .nsmap[["c"]])
      if (!is.null(val)) xml2::xml_set_attr(nd, "val", as.character(val))
      wrap_element(nd)
    },
    .remove_majorTickMark = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:majorTickMark", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .add_minorTickMark = function(val = NULL) {
      nd <- xml2::xml_add_child(self$get_node(), "c:minorTickMark", xmlns = .nsmap[["c"]])
      if (!is.null(val)) xml2::xml_set_attr(nd, "val", as.character(val))
      wrap_element(nd)
    },
    .remove_minorTickMark = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:minorTickMark", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    get_or_add_scaling = function() {
      nd <- .get_or_add_child(self$get_node(), "c:scaling", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_delete_ = function() {
      nd <- .get_or_add_child(self$get_node(), "c:delete", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_spPr = function() {
      .get_or_add_child(self$get_node(), "c:spPr", .nsmap[["c"]])
    }
  ),

  active = list(
    scaling = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:scaling", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    title = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:title", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    majorGridlines = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:majorGridlines", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    minorGridlines = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:minorGridlines", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    majorTickMark = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:majorTickMark", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    minorTickMark = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:minorTickMark", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    tickLblPos = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:tickLblPos", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    numFmt = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:numFmt", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    delete_ = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:delete", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    spPr = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:spPr", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    orientation = function(value) {
      sc <- self$scaling
      if (!missing(value)) {
        if (is.null(sc)) sc <- self$get_or_add_scaling()
        sc$.remove_orientation()
        if (value == "maxMin") {
          orient_elm <- sc$get_or_add_orientation()
          orient_elm$val <- value
        }
        return(invisible(value))
      }
      if (is.null(sc)) return("minMax")
      orient <- sc$orientation
      if (is.null(orient)) return("minMax")
      orient$val
    },

    # defRPr: txPr > a:p > a:pPr > a:defRPr (chain created on demand)
    defRPr = function() {
      txPr <- self$get_or_add_txPr()
      .defRPr_from_txPr(txPr)
    },

    # c:lblOffset child (catAx only but defined here for convenience)
    lblOffset = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:lblOffset", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    # c:crossAx child
    crossAx = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:crossAx", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    # c:crosses child
    crosses = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:crosses", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    # c:crossesAt child
    crossesAt = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:crossesAt", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    # c:majorUnit child (valAx only)
    majorUnit = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:majorUnit", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    # c:minorUnit child (valAx only)
    minorUnit = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:minorUnit", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  ),

  private = list()
)


# Helper: get defRPr from a txPr node (c: or a: namespace)
.defRPr_from_txPr <- function(txPr_node) {
  c_ns <- .nsmap[["c"]]
  a_ns <- .nsmap[["a"]]
  # txPr is a CT_TextBody; navigate to a:p/a:pPr/a:defRPr
  p_nd <- xml2::xml_find_first(txPr_node, "a:p", ns = c(a = a_ns))
  if (inherits(p_nd, "xml_missing")) {
    p_nd <- xml2::xml_add_child(txPr_node, xml2::read_xml(
      sprintf('<a:p xmlns:a="%s"/>', a_ns)
    ))
  }
  pPr_nd <- xml2::xml_find_first(p_nd, "a:pPr", ns = c(a = a_ns))
  if (inherits(pPr_nd, "xml_missing")) {
    pPr_nd <- xml2::xml_add_child(p_nd, xml2::read_xml(
      sprintf('<a:pPr xmlns:a="%s"/>', a_ns)
    ))
  }
  defRPr_nd <- xml2::xml_find_first(pPr_nd, "a:defRPr", ns = c(a = a_ns))
  if (inherits(defRPr_nd, "xml_missing")) {
    defRPr_nd <- xml2::xml_add_child(pPr_nd, xml2::read_xml(
      sprintf('<a:defRPr xmlns:a="%s"/>', a_ns)
    ))
  }
  wrap_element(defRPr_nd)
}

# Helper: get or add a direct child by qualified name
.get_or_add_child <- function(parent_node, qname, ns_uri) {
  prefix <- sub(":.*", "", qname)
  local  <- sub(".*:", "", qname)
  ns_vec <- setNames(ns_uri, prefix)
  nd <- xml2::xml_find_first(parent_node, qname, ns = ns_vec)
  if (inherits(nd, "xml_missing")) {
    nd <- xml2::xml_add_child(parent_node, qname, xmlns = ns_uri)
  }
  nd
}


# ============================================================================
# CT_CatAx — <c:catAx>
# ============================================================================

#' Category axis XML element
#' @noRd
#' @export
CT_CatAx <- R6::R6Class(
  "CT_CatAx",
  inherit = BaseAxisElement,
  public = list(
    .add_lblOffset = function() {
      nd <- xml2::xml_add_child(self$get_node(), "c:lblOffset", xmlns = .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_lblOffset = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:lblOffset", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .remove_crosses = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:crosses", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .remove_crossesAt = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:crossesAt", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .add_crosses = function(val = NULL) {
      nd <- xml2::xml_add_child(self$get_node(), "c:crosses", xmlns = .nsmap[["c"]])
      if (!is.null(val)) xml2::xml_set_attr(nd, "val", as.character(val))
      wrap_element(nd)
    },
    .add_crossesAt = function(val = NULL) {
      nd <- xml2::xml_add_child(self$get_node(), "c:crossesAt", xmlns = .nsmap[["c"]])
      if (!is.null(val)) xml2::xml_set_attr(nd, "val", as.character(as.numeric(val)))
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_DateAx — <c:dateAx>
# ============================================================================

#' Date axis XML element
#' @noRd
#' @export
CT_DateAx <- R6::R6Class(
  "CT_DateAx",
  inherit = BaseAxisElement
)


# ============================================================================
# CT_ValAx — <c:valAx>
# ============================================================================

#' Value axis XML element
#' @noRd
#' @export
CT_ValAx <- R6::R6Class(
  "CT_ValAx",
  inherit = BaseAxisElement,
  public = list(
    .remove_crosses = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:crosses", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .remove_crossesAt = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:crossesAt", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .add_crosses = function(val = NULL) {
      nd <- xml2::xml_add_child(self$get_node(), "c:crosses", xmlns = .nsmap[["c"]])
      if (!is.null(val)) xml2::xml_set_attr(nd, "val", as.character(val))
      wrap_element(nd)
    },
    .add_crossesAt = function(val = NULL) {
      nd <- xml2::xml_add_child(self$get_node(), "c:crossesAt", xmlns = .nsmap[["c"]])
      if (!is.null(val)) xml2::xml_set_attr(nd, "val", as.character(as.numeric(val)))
      wrap_element(nd)
    },
    .add_majorUnit = function(val = NULL) {
      nd <- xml2::xml_add_child(self$get_node(), "c:majorUnit", xmlns = .nsmap[["c"]])
      if (!is.null(val)) xml2::xml_set_attr(nd, "val", as.character(as.numeric(val)))
      wrap_element(nd)
    },
    .remove_majorUnit = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:majorUnit", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .add_minorUnit = function(val = NULL) {
      nd <- xml2::xml_add_child(self$get_node(), "c:minorUnit", xmlns = .nsmap[["c"]])
      if (!is.null(val)) xml2::xml_set_attr(nd, "val", as.character(as.numeric(val)))
      wrap_element(nd)
    },
    .remove_minorUnit = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:minorUnit", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    }
  )
)


# ============================================================================
# CT_ChartLines — <c:majorGridlines>, <c:minorGridlines>
# ============================================================================

#' Chart gridlines XML element
#' @noRd
#' @export
CT_ChartLines <- R6::R6Class(
  "CT_ChartLines",
  inherit = BaseOxmlElement,
  public = list(
    get_or_add_spPr = function() {
      .get_or_add_child(self$get_node(), "c:spPr", .nsmap[["c"]])
    }
  )
)


# ============================================================================
# CT_Tx — <c:tx>
# ============================================================================

#' Chart text element
#' @noRd
#' @export
CT_Tx <- R6::R6Class(
  "CT_Tx",
  inherit = BaseOxmlElement,

  public = list(
    get_or_add_rich = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:rich", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) {
        xml_str <- sprintf(
          '<c:rich xmlns:c="%s" xmlns:a="%s"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:rich>',
          .nsmap[["c"]], .nsmap[["a"]]
        )
        nd_doc <- xml2::read_xml(xml_str)
        nd <- xml2::xml_root(nd_doc)
        xml2::xml_add_child(self$get_node(), nd)
        nd <- xml2::xml_find_first(self$get_node(), "c:rich", ns = c(c = .nsmap[["c"]]))
      }
      wrap_element(nd)
    },
    .remove_strRef = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:strRef", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    }
  ),

  active = list(
    rich = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:rich", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    strRef = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:strRef", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_Title — <c:title>
# ============================================================================

#' Chart title element
#' @noRd
#' @export
CT_Title <- R6::R6Class(
  "CT_Title",
  inherit = BaseOxmlElement,

  public = list(
    get_or_add_tx = function() {
      nd <- .get_or_add_child(self$get_node(), "c:tx", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_tx_rich = function() {
      tx <- self$get_or_add_tx()
      tx$.remove_strRef()
      tx$get_or_add_rich()
      # Return the tx element (the rich child is inside it)
      tx
    },
    .remove_tx = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:tx", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    get_or_add_spPr = function() {
      .get_or_add_child(self$get_node(), "c:spPr", .nsmap[["c"]])
    }
  ),

  active = list(
    tx = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:tx", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    tx_rich = function() {
      nds <- self$xpath("c:tx/c:rich")
      if (length(nds) == 0) return(NULL)
      wrap_element(nds[[1]])
    },
    spPr = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:spPr", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_LegendPos — <c:legendPos val="...">
# ============================================================================

#' Legend position element
#' @noRd
#' @export
CT_LegendPos <- R6::R6Class(
  "CT_LegendPos",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(value))
        return(invisible(value))
      }
      v <- self$get_attr("val")
      if (is.null(v) || is.na(v)) return(XL_LEGEND_POSITION$RIGHT)
      v
    }
  )
)


# ============================================================================
# CT_Legend — <c:legend>
# ============================================================================

#' Chart legend XML element
#' @noRd
#' @export
CT_Legend <- R6::R6Class(
  "CT_Legend",
  inherit = BaseOxmlElement,

  public = list(
    get_or_add_legendPos = function() {
      nd <- .get_or_add_child(self$get_node(), "c:legendPos", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_overlay = function() {
      nd <- .get_or_add_child(self$get_node(), "c:overlay", .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_overlay = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:overlay", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    get_or_add_layout = function() {
      nd <- .get_or_add_child(self$get_node(), "c:layout", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_txPr = function() {
      nd <- .get_or_add_child(self$get_node(), "c:txPr", .nsmap[["c"]])
      nd
    }
  ),

  active = list(
    legendPos = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:legendPos", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    overlay = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:overlay", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    layout = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:layout", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    defRPr = function() {
      txPr <- self$get_or_add_txPr()
      .defRPr_from_txPr(txPr)
    },
    horz_offset = function(value) {
      layout <- self$layout
      if (!missing(value)) {
        if (value == 0.0) {
          nd <- xml2::xml_find_first(self$get_node(), "c:layout", ns = c(c = .nsmap[["c"]]))
          if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
          return(invisible(value))
        }
        layout <- self$get_or_add_layout()
        layout$horz_offset <- value
        return(invisible(value))
      }
      if (is.null(layout)) return(0.0)
      layout$horz_offset
    }
  )
)


# ============================================================================
# CT_Layout, CT_ManualLayout — <c:layout>, <c:manualLayout>
# ============================================================================

#' Chart layout XML element
#' @noRd
#' @export
CT_Layout <- R6::R6Class(
  "CT_Layout",
  inherit = BaseOxmlElement,
  public = list(
    get_or_add_manualLayout = function() {
      nd <- .get_or_add_child(self$get_node(), "c:manualLayout", .nsmap[["c"]])
      wrap_element(nd)
    }
  ),
  active = list(
    manualLayout = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:manualLayout", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    horz_offset = function(value) {
      ml <- self$manualLayout
      if (!missing(value)) {
        if (value == 0.0) {
          nd <- xml2::xml_find_first(self$get_node(), "c:manualLayout", ns = c(c = .nsmap[["c"]]))
          if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
          return(invisible(value))
        }
        ml <- self$get_or_add_manualLayout()
        ml$horz_offset <- value
        return(invisible(value))
      }
      if (is.null(ml)) return(0.0)
      ml$horz_offset
    }
  )
)

#' Chart manual layout XML element
#' @noRd
#' @export
CT_ManualLayout <- R6::R6Class(
  "CT_ManualLayout",
  inherit = BaseOxmlElement,
  public = list(
    get_or_add_xMode = function() {
      nd <- .get_or_add_child(self$get_node(), "c:xMode", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_x = function() {
      nd <- .get_or_add_child(self$get_node(), "c:x", .nsmap[["c"]])
      wrap_element(nd)
    }
  ),
  active = list(
    xMode = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:xMode", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    x = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:x", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    horz_offset = function(value) {
      if (!missing(value)) {
        xm <- self$get_or_add_xMode(); xm$val <- "factor"
        xv <- self$get_or_add_x();     xv$val <- value
        return(invisible(value))
      }
      x     <- self$x
      xMode <- self$xMode
      if (is.null(x) || is.null(xMode) || xMode$val != "factor") return(0.0)
      x$val
    }
  )
)

#' XML element wrapper for c:layoutMode
#' @noRd
#' @export
CT_LayoutMode <- R6::R6Class(
  "CT_LayoutMode",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(value))
        return(invisible(value))
      }
      v <- self$get_attr("val")
      if (is.null(v) || is.na(v)) return("factor")
      v
    }
  )
)


# ============================================================================
# CT_DLblPos — <c:dLblPos val="...">
# ============================================================================

#' Data label position element
#' @noRd
#' @export
CT_DLblPos <- R6::R6Class(
  "CT_DLblPos",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(value))
        return(invisible(value))
      }
      self$get_attr("val")
    }
  )
)


# ============================================================================
# CT_DLbls — <c:dLbls>
# ============================================================================

#' Data labels container element
#' @noRd
#' @export
CT_DLbls <- R6::R6Class(
  "CT_DLbls",
  inherit = BaseOxmlElement,

  public = list(
    get_or_add_txPr = function() {
      nd <- .get_or_add_child(self$get_node(), "c:txPr", .nsmap[["c"]])
      nd
    },
    get_or_add_numFmt = function() {
      nd <- .get_or_add_child(self$get_node(), "c:numFmt", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_dLblPos = function() {
      nd <- .get_or_add_child(self$get_node(), "c:dLblPos", .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_dLblPos = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:dLblPos", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    get_or_add_showCatName = function() {
      nd <- .get_or_add_child(self$get_node(), "c:showCatName", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_showLegendKey = function() {
      nd <- .get_or_add_child(self$get_node(), "c:showLegendKey", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_showPercent = function() {
      nd <- .get_or_add_child(self$get_node(), "c:showPercent", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_showSerName = function() {
      nd <- .get_or_add_child(self$get_node(), "c:showSerName", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_showVal = function() {
      nd <- .get_or_add_child(self$get_node(), "c:showVal", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_dLbl_for_point = function(idx) {
      nds <- self$xpath(sprintf('c:dLbl[c:idx[@val="%d"]]', idx))
      if (length(nds) == 0) return(NULL)
      wrap_element(nds[[1]])
    },
    get_or_add_dLbl_for_point = function(idx) {
      existing <- self$get_dLbl_for_point(idx)
      if (!is.null(existing)) return(existing)
      self$.insert_dLbl_in_sequence(idx)
    },
    .insert_dLbl_in_sequence = function(idx) {
      new_dLbl <- .new_dLbl_element()
      # set idx value
      idx_nd <- xml2::xml_find_first(new_dLbl, "c:idx", ns = c(c = .nsmap[["c"]]))
      xml2::xml_set_attr(idx_nd, "val", as.character(as.integer(idx)))
      # find insertion point
      existing_dLbls <- self$xpath("c:dLbl")
      inserted <- FALSE
      for (dl in existing_dLbls) {
        dl_idx_nd <- xml2::xml_find_first(dl, "c:idx", ns = c(c = .nsmap[["c"]]))
        dl_idx_val <- as.integer(xml2::xml_attr(dl_idx_nd, "val"))
        if (dl_idx_val > idx) {
          xml2::xml_add_sibling(dl, new_dLbl, .where = "before")
          inserted <- TRUE
          break
        }
      }
      if (!inserted) {
        xml2::xml_add_child(self$get_node(), new_dLbl)
      }
      # return wrapped newly inserted element
      self$get_dLbl_for_point(idx)
    }
  ),

  active = list(
    numFmt = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:numFmt", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    dLblPos = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:dLblPos", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    defRPr = function() {
      txPr <- self$get_or_add_txPr()
      .defRPr_from_txPr(txPr)
    }
  )
)

.new_dLbl_element <- function() {
  xml_str <- sprintf(
    '<c:dLbl xmlns:c="%s" xmlns:a="%s">
  <c:idx val="666"/>
  <c:spPr/>
  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr>
  <c:showLegendKey val="0"/>
  <c:showVal val="1"/>
  <c:showCatName val="0"/>
  <c:showSerName val="0"/>
  <c:showPercent val="0"/>
  <c:showBubbleSize val="0"/>
</c:dLbl>',
    .nsmap[["c"]], .nsmap[["a"]]
  )
  xml2::xml_root(xml2::read_xml(xml_str))
}


# ============================================================================
# CT_DLbl — <c:dLbl>
# ============================================================================

#' Individual data label element
#' @noRd
#' @export
CT_DLbl <- R6::R6Class(
  "CT_DLbl",
  inherit = BaseOxmlElement,

  public = list(
    get_or_add_tx = function() {
      nd <- .get_or_add_child(self$get_node(), "c:tx", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_txPr = function() {
      .get_or_add_child(self$get_node(), "c:txPr", .nsmap[["c"]])
    },
    get_or_add_dLblPos = function() {
      nd <- .get_or_add_child(self$get_node(), "c:dLblPos", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_rich = function() {
      tx <- self$get_or_add_tx()
      tx$.remove_strRef()
      tx$get_or_add_rich()
    },
    get_or_add_tx_rich = function() {
      tx <- self$get_or_add_tx()
      tx$.remove_strRef()
      tx$get_or_add_rich()
      tx
    },
    remove_tx_rich = function() {
      nds <- self$xpath("c:tx[c:rich]")
      if (length(nds) > 0) xml2::xml_remove(nds[[1]])
    },
    .remove_spPr = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:spPr", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .remove_txPr = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:txPr", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .remove_dLblPos = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:dLblPos", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    }
  ),

  active = list(
    idx = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:idx", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    idx_val = function() {
      idx <- self$idx
      if (is.null(idx)) return(NULL)
      idx$val
    },
    tx = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:tx", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    dLblPos = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:dLblPos", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_Marker — <c:marker>
# ============================================================================

#' Marker element for line/xy/radar charts
#' @noRd
#' @export
CT_Marker <- R6::R6Class(
  "CT_Marker",
  inherit = BaseOxmlElement,

  public = list(
    get_or_add_spPr = function() {
      .get_or_add_child(self$get_node(), "c:spPr", .nsmap[["c"]])
    },
    .add_symbol = function() {
      nd <- xml2::xml_add_child(self$get_node(), "c:symbol", xmlns = .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_symbol = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:symbol", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .add_size = function() {
      nd <- xml2::xml_add_child(self$get_node(), "c:size", xmlns = .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_size = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:size", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    }
  ),

  active = list(
    symbol = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:symbol", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    symbol_val = function() {
      s <- self$symbol
      if (is.null(s)) return(NULL)
      s$val
    },
    size = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:size", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    size_val = function() {
      sz <- self$size
      if (is.null(sz)) return(NULL)
      sz$val
    },
    spPr = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:spPr", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)

#' Marker style element
#' @noRd
#' @export
CT_MarkerStyle <- R6::R6Class(
  "CT_MarkerStyle",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(value))
        return(invisible(value))
      }
      self$get_attr("val")
    }
  )
)

#' Marker size element
#' @noRd
#' @export
CT_MarkerSize <- R6::R6Class(
  "CT_MarkerSize",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(as.integer(value)))
        return(invisible(value))
      }
      as.integer(self$get_attr("val"))
    }
  )
)


# ============================================================================
# CT_DPt — <c:dPt>
# ============================================================================

#' Data point element
#' @noRd
#' @export
CT_DPt <- R6::R6Class(
  "CT_DPt",
  inherit = BaseOxmlElement,
  public = list(
    get_or_add_spPr = function() {
      .get_or_add_child(self$get_node(), "c:spPr", .nsmap[["c"]])
    },
    get_or_add_marker = function() {
      nd <- .get_or_add_child(self$get_node(), "c:marker", .nsmap[["c"]])
      wrap_element(nd)
    }
  ),
  active = list(
    idx = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:idx", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    marker = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:marker", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    spPr = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:spPr", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_SeriesComposite — <c:ser>
# ============================================================================

#' Chart series element (composite; all series use the same c:ser tag)
#' @noRd
#' @export
CT_SeriesComposite <- R6::R6Class(
  "CT_SeriesComposite",
  inherit = BaseOxmlElement,

  public = list(
    get_or_add_dLbls = function() {
      nd <- .get_or_add_child(self$get_node(), "c:dLbls", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_marker = function() {
      nd <- .get_or_add_child(self$get_node(), "c:marker", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_invertIfNegative = function() {
      nd <- .get_or_add_child(self$get_node(), "c:invertIfNegative", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_smooth = function() {
      nd <- .get_or_add_child(self$get_node(), "c:smooth", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_spPr = function() {
      .get_or_add_child(self$get_node(), "c:spPr", .nsmap[["c"]])
    },
    get_dLbl = function(idx) {
      dl <- self$dLbls
      if (is.null(dl)) return(NULL)
      dl$get_dLbl_for_point(idx)
    },
    get_or_add_dLbl = function(idx) {
      self$get_or_add_dLbls()$get_or_add_dLbl_for_point(idx)
    },
    get_or_add_dPt_for_point = function(idx) {
      nds <- self$xpath(sprintf('c:dPt[c:idx[@val="%d"]]', idx))
      if (length(nds) > 0) return(wrap_element(nds[[1]]))
      dPt_xml <- sprintf(
        '<c:dPt xmlns:c="%s"><c:idx val="%d"/></c:dPt>',
        .nsmap[["c"]], as.integer(idx)
      )
      dPt_nd <- xml2::xml_root(xml2::read_xml(dPt_xml))
      xml2::xml_add_child(self$get_node(), dPt_nd)
      nd <- self$xpath(sprintf('c:dPt[c:idx[@val="%d"]]', idx))[[1]]
      wrap_element(nd)
    }
  ),

  active = list(
    idx = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:idx", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    order = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:order", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    dLbls = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:dLbls", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    marker = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:marker", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    invertIfNegative = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:invertIfNegative", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    smooth = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:smooth", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    spPr = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:spPr", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    val = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:val", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    yVal = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:yVal", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    xVal = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:xVal", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    bubbleSize = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:bubbleSize", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    cat = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:cat", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    cat_ptCount_val = function() {
      vals <- self$xpath("./c:cat//c:ptCount/@val")
      if (length(vals) == 0) return(0L)
      as.integer(xml2::xml_text(vals[[1]]))
    },
    xVal_ptCount_val = function() {
      vals <- self$xpath("./c:xVal//c:ptCount/@val")
      if (length(vals) == 0) return(0L)
      as.integer(xml2::xml_text(vals[[1]]))
    },
    yVal_ptCount_val = function() {
      vals <- self$xpath("./c:yVal//c:ptCount/@val")
      if (length(vals) == 0) return(0L)
      as.integer(xml2::xml_text(vals[[1]]))
    },
    bubbleSize_ptCount_val = function() {
      vals <- self$xpath("./c:bubbleSize//c:ptCount/@val")
      if (length(vals) == 0) return(0L)
      as.integer(xml2::xml_text(vals[[1]]))
    }
  )
)


# ============================================================================
# CT_NumDataSource — <c:yVal>, <c:val> etc.
# ============================================================================

#' Numeric data source element
#' @noRd
#' @export
CT_NumDataSource <- R6::R6Class(
  "CT_NumDataSource",
  inherit = BaseOxmlElement,
  active = list(
    ptCount_val = function() {
      nds <- self$xpath(".//c:ptCount/@val")
      if (length(nds) == 0) return(0L)
      as.integer(xml2::xml_text(nds[[1]]))
    }
  ),
  public = list(
    pt_v = function(idx) {
      nds <- self$xpath(sprintf(".//c:pt[@idx=%d]", idx))
      if (length(nds) == 0) return(NULL)
      v_nd <- xml2::xml_find_first(nds[[1]], "c:v", ns = c(c = .nsmap[["c"]]))
      if (inherits(v_nd, "xml_missing")) return(NULL)
      as.numeric(xml2::xml_text(v_nd))
    }
  )
)


# ============================================================================
# CT_PlotArea — <c:plotArea>
# ============================================================================

.xChart_tags <- c(
  paste0("{", .nsmap[["c"]], "}area3DChart"),
  paste0("{", .nsmap[["c"]], "}areaChart"),
  paste0("{", .nsmap[["c"]], "}bar3DChart"),
  paste0("{", .nsmap[["c"]], "}barChart"),
  paste0("{", .nsmap[["c"]], "}bubbleChart"),
  paste0("{", .nsmap[["c"]], "}doughnutChart"),
  paste0("{", .nsmap[["c"]], "}line3DChart"),
  paste0("{", .nsmap[["c"]], "}lineChart"),
  paste0("{", .nsmap[["c"]], "}ofPieChart"),
  paste0("{", .nsmap[["c"]], "}pie3DChart"),
  paste0("{", .nsmap[["c"]], "}pieChart"),
  paste0("{", .nsmap[["c"]], "}radarChart"),
  paste0("{", .nsmap[["c"]], "}scatterChart"),
  paste0("{", .nsmap[["c"]], "}stockChart"),
  paste0("{", .nsmap[["c"]], "}surface3DChart"),
  paste0("{", .nsmap[["c"]], "}surfaceChart")
)

#' Plot area XML element
#' @noRd
#' @export
CT_PlotArea <- R6::R6Class(
  "CT_PlotArea",
  inherit = BaseOxmlElement,

  public = list(
    iter_xCharts = function() {
      children <- xml2::xml_children(self$get_node())
      Filter(function(nd) xml2::xml_name(nd, ns = xml2::xml_ns(nd)) %in% {
        c("c:area3DChart","c:areaChart","c:bar3DChart","c:barChart",
          "c:bubbleChart","c:doughnutChart","c:line3DChart","c:lineChart",
          "c:ofPieChart","c:pie3DChart","c:pieChart","c:radarChart",
          "c:scatterChart","c:stockChart","c:surface3DChart","c:surfaceChart")
      }, children)
    },

    iter_sers = function() {
      result <- list()
      for (xc in self$iter_xCharts()) {
        sers_in_xc <- xml2::xml_find_all(xc, "c:ser", ns = c(c = .nsmap[["c"]]))
        # sort by c:order/@val
        orders <- sapply(sers_in_xc, function(s) {
          ord_nd <- xml2::xml_find_first(s, "c:order", ns = c(c = .nsmap[["c"]]))
          if (inherits(ord_nd, "xml_missing")) return(0L)
          as.integer(xml2::xml_attr(ord_nd, "val"))
        })
        sers_in_xc <- sers_in_xc[order(orders)]
        result <- c(result, lapply(sers_in_xc, wrap_element))
      }
      result
    }
  ),

  active = list(
    xCharts = function() {
      lapply(self$iter_xCharts(), wrap_element)
    },

    sers = function() {
      self$iter_sers()
    },

    catAx_lst = function() {
      nds <- xml2::xml_find_all(self$get_node(), "c:catAx", ns = c(c = .nsmap[["c"]]))
      lapply(nds, wrap_element)
    },

    valAx_lst = function() {
      nds <- xml2::xml_find_all(self$get_node(), "c:valAx", ns = c(c = .nsmap[["c"]]))
      lapply(nds, wrap_element)
    },

    dateAx_lst = function() {
      nds <- xml2::xml_find_all(self$get_node(), "c:dateAx", ns = c(c = .nsmap[["c"]]))
      lapply(nds, wrap_element)
    }
  )
)


# ============================================================================
# CT_Chart — <c:chart>
# ============================================================================

#' Chart element
#' @noRd
#' @export
CT_Chart <- R6::R6Class(
  "CT_Chart",
  inherit = BaseOxmlElement,

  public = list(
    get_or_add_title = function() {
      nd <- .get_or_add_child(self$get_node(), "c:title", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_autoTitleDeleted = function() {
      nd <- .get_or_add_child(self$get_node(), "c:autoTitleDeleted", .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_title = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:title", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .add_legend = function() {
      nd <- xml2::xml_add_child(self$get_node(), "c:legend", xmlns = .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_legend = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:legend", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    }
  ),

  active = list(
    title = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:title", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    plotArea = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:plotArea", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    legend = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:legend", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    has_legend = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) {
          if (is.null(self$legend)) self$.add_legend()
        } else {
          self$.remove_legend()
        }
        return(invisible(value))
      }
      !is.null(self$legend)
    }
  )
)


# ============================================================================
# CT_ChartSpace — <c:chartSpace> (root element of chart part)
# ============================================================================

#' Chart space XML element (root of chart part)
#' @noRd
#' @export
CT_ChartSpace <- R6::R6Class(
  "CT_ChartSpace",
  inherit = BaseOxmlElement,

  public = list(
    get_or_add_txPr = function() {
      nd <- .get_or_add_child(self$get_node(), "c:txPr", .nsmap[["c"]])
      nd
    },
    .remove_style = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:style", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    .add_style = function(val = NULL) {
      nd <- xml2::xml_add_child(self$get_node(), "c:style", xmlns = .nsmap[["c"]])
      if (!is.null(val)) xml2::xml_set_attr(nd, "val", as.character(as.integer(val)))
      wrap_element(nd)
    },
    get_or_add_title = function() {
      self$chart$get_or_add_title()
    }
  ),

  active = list(
    chart = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:chart", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    plotArea = function() {
      self$chart$plotArea
    },
    style = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:style", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },
    catAx_lst = function() self$chart$plotArea$catAx_lst,
    dateAx_lst = function() self$chart$plotArea$dateAx_lst,
    valAx_lst  = function() self$chart$plotArea$valAx_lst
  )
)


# ============================================================================
# CT_Style — <c:style val="...">
# ============================================================================

#' Chart style element
#' @noRd
#' @export
CT_Style <- R6::R6Class(
  "CT_Style",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(as.integer(value)))
        return(invisible(value))
      }
      as.integer(self$get_attr("val"))
    }
  )
)


# ============================================================================
# Plot/xChart element — generic base for c:barChart, c:lineChart etc.
# ============================================================================

#' Generic xChart element (barChart, lineChart, etc.)
#' @noRd
#' @export
CT_xChart <- R6::R6Class(
  "CT_xChart",
  inherit = BaseOxmlElement,

  public = list(
    iter_sers = function() {
      sers_nds <- xml2::xml_find_all(self$get_node(), "c:ser", ns = c(c = .nsmap[["c"]]))
      orders <- sapply(sers_nds, function(s) {
        ord_nd <- xml2::xml_find_first(s, "c:order", ns = c(c = .nsmap[["c"]]))
        if (inherits(ord_nd, "xml_missing")) return(0L)
        as.integer(xml2::xml_attr(ord_nd, "val"))
      })
      sers_nds <- sers_nds[order(orders)]
      lapply(sers_nds, wrap_element)
    },
    .remove_dLbls = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:dLbls", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    },
    add_dLbls_default = function() {
      nd_xml <- sprintf(
        '<c:dLbls xmlns:c="%s"><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>',
        .nsmap[["c"]]
      )
      nd <- xml2::xml_root(xml2::read_xml(nd_xml))
      xml2::xml_add_child(self$get_node(), nd)
      nd2 <- xml2::xml_find_first(self$get_node(), "c:dLbls", ns = c(c = .nsmap[["c"]]))
      wrap_element(nd2)
    },
    get_or_add_varyColors = function() {
      nd <- .get_or_add_child(self$get_node(), "c:varyColors", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_gapWidth = function() {
      nd <- .get_or_add_child(self$get_node(), "c:gapWidth", .nsmap[["c"]])
      wrap_element(nd)
    },
    get_or_add_overlap = function() {
      nd <- .get_or_add_child(self$get_node(), "c:overlap", .nsmap[["c"]])
      wrap_element(nd)
    },
    .remove_overlap = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:overlap", ns = c(c = .nsmap[["c"]]))
      if (!inherits(nd, "xml_missing")) xml2::xml_remove(nd)
    }
  ),

  active = list(
    sers = function() self$iter_sers(),

    dLbls = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:dLbls", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    varyColors = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:varyColors", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    # barChart / bar3DChart
    barDir = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:barDir", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    grouping_val = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:grouping", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      xml2::xml_attr(nd, "val")
    },

    gapWidth = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:gapWidth", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    overlap = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:overlap", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    },

    # bubbleChart
    bubbleScale = function() {
      nd <- xml2::xml_find_first(self$get_node(), "c:bubbleScale", ns = c(c = .nsmap[["c"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_GapAmount / CT_Overlap — simple integer elements
# ============================================================================

#' Gap width / overlap integer element
#' @noRd
#' @export
CT_IntegerElement <- R6::R6Class(
  "CT_IntegerElement",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(as.integer(value)))
        return(invisible(value))
      }
      as.integer(self$get_attr("val"))
    }
  )
)


# ============================================================================
# CT_BarDir — <c:barDir val="bar|col">
# ============================================================================

#' Bar direction element
#' @noRd
#' @export
CT_BarDir <- R6::R6Class(
  "CT_BarDir",
  inherit = BaseOxmlElement,
  active = list(
    val = function(value) {
      if (!missing(value)) {
        self$set_attr("val", as.character(value))
        return(invisible(value))
      }
      self$get_attr("val")
    }
  )
)


# ============================================================================
# .onLoad registration
# ============================================================================

.onLoad_oxml_chart <- function() {
  register_element_cls("c:chartSpace",  CT_ChartSpace)
  register_element_cls("c:chart",       CT_Chart)
  register_element_cls("c:plotArea",    CT_PlotArea)
  register_element_cls("c:title",       CT_Title)
  register_element_cls("c:tx",          CT_Tx)
  register_element_cls("c:legend",      CT_Legend)
  register_element_cls("c:legendPos",   CT_LegendPos)
  register_element_cls("c:layout",      CT_Layout)
  register_element_cls("c:manualLayout", CT_ManualLayout)
  register_element_cls("c:xMode",       CT_LayoutMode)
  register_element_cls("c:x",           CT_Double)
  register_element_cls("c:catAx",       CT_CatAx)
  register_element_cls("c:valAx",       CT_ValAx)
  register_element_cls("c:dateAx",      CT_DateAx)
  register_element_cls("c:scaling",     CT_Scaling)
  register_element_cls("c:orientation", CT_Orientation)
  register_element_cls("c:lblOffset",   CT_LblOffset)
  register_element_cls("c:tickLblPos",  CT_TickLblPos)
  register_element_cls("c:majorTickMark", CT_TickMark)
  register_element_cls("c:minorTickMark", CT_TickMark)
  register_element_cls("c:crosses",     CT_Crosses)
  register_element_cls("c:crossesAt",   CT_CrossesAt)
  register_element_cls("c:numFmt",      CT_NumFmt)
  register_element_cls("c:ser",         CT_SeriesComposite)
  register_element_cls("c:dLbls",       CT_DLbls)
  register_element_cls("c:dLbl",        CT_DLbl)
  register_element_cls("c:dLblPos",     CT_DLblPos)
  register_element_cls("c:marker",      CT_Marker)
  register_element_cls("c:symbol",      CT_MarkerStyle)
  register_element_cls("c:size",        CT_MarkerSize)
  register_element_cls("c:dPt",         CT_DPt)
  register_element_cls("c:idx",         CT_UnsignedInt)
  register_element_cls("c:order",       CT_UnsignedInt)
  register_element_cls("c:majorGridlines", CT_ChartLines)
  register_element_cls("c:minorGridlines", CT_ChartLines)
  register_element_cls("c:style",       CT_Style)
  register_element_cls("c:delete",      CT_Boolean)
  register_element_cls("c:overlay",     CT_Boolean)
  register_element_cls("c:autoTitleDeleted", CT_Boolean)
  register_element_cls("c:showVal",     CT_Boolean)
  register_element_cls("c:showCatName", CT_Boolean)
  register_element_cls("c:showSerName", CT_Boolean)
  register_element_cls("c:showPercent", CT_Boolean)
  register_element_cls("c:showBubbleSize", CT_Boolean)
  register_element_cls("c:showLegendKey",  CT_Boolean)
  register_element_cls("c:invertIfNegative", CT_Boolean)
  register_element_cls("c:smooth",      CT_Boolean)
  register_element_cls("c:varyColors",  CT_Boolean)
  register_element_cls("c:majorUnit",   CT_AxisUnit)
  register_element_cls("c:minorUnit",   CT_AxisUnit)
  register_element_cls("c:max",         CT_Double)
  register_element_cls("c:min",         CT_Double)
  register_element_cls("c:barDir",      CT_BarDir)
  register_element_cls("c:gapWidth",    CT_IntegerElement)
  register_element_cls("c:overlap",     CT_IntegerElement)
  register_element_cls("c:val",         CT_NumDataSource)
  register_element_cls("c:yVal",        CT_NumDataSource)
  register_element_cls("c:xVal",        CT_NumDataSource)
  # xChart elements all use CT_xChart
  for (tag in c("c:areaChart","c:area3DChart","c:barChart","c:bar3DChart",
                "c:bubbleChart","c:doughnutChart","c:lineChart","c:line3DChart",
                "c:pieChart","c:pie3DChart","c:radarChart","c:scatterChart",
                "c:stockChart","c:surfaceChart","c:surface3DChart","c:ofPieChart")) {
    register_element_cls(tag, CT_xChart)
  }
}
