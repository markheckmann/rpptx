# DML fill domain objects.
#
# Ported from python-pptx/src/pptx/dml/fill.py.
# Provides FillFormat R6 class and gradient stop helpers.

# ============================================================================
# Helpers — locate and remove fill elements within a shape-properties element
# ============================================================================

# Return the first fill-choice child of an spPr/ln element, or NULL
.get_fill_elm <- function(spPr) {
  for (child in xml2::xml_children(spPr$get_node())) {
    if (.get_clark_name(child) %in% .fill_choice_tags) {
      return(wrap_element(child))
    }
  }
  NULL
}

# Remove all fill-choice children from an spPr/ln element
.remove_fill_elms <- function(spPr) {
  node <- spPr$get_node()
  for (child in xml2::xml_children(node)) {
    if (.get_clark_name(child) %in% .fill_choice_tags) {
      xml2::xml_remove(child)
    }
  }
  invisible(NULL)
}

# Insert an XML fragment as a fill child (before ln/effectLst/extLst)
.insert_fill_child <- function(spPr, xml_str) {
  a <- .nsmap[["a"]]
  p <- .nsmap[["p"]]
  # Successors: ln, effectLst, effectDag, scene3d, sp3d, extLst
  successors <- c(qn("a:ln"), qn("a:effectLst"), qn("a:effectDag"),
                  qn("a:scene3d"), qn("a:sp3d"), qn("a:extLst"))
  new_elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
  node <- spPr$get_node()
  inserted <- FALSE
  for (child in xml2::xml_children(node)) {
    if (.get_clark_name(child) %in% successors) {
      xml2::xml_add_sibling(child, new_elm$get_node(), .where = "before")
      inserted <- TRUE
      break
    }
  }
  if (!inserted) {
    xml2::xml_add_child(node, new_elm$get_node())
  }
  # Return wrapped version of the newly-inserted node
  .get_fill_elm(spPr)
}


# ============================================================================
# GradientStop — wraps a <a:gs> element
# ============================================================================

#' Gradient stop proxy
#' @noRd
GradientStop <- R6::R6Class(
  "GradientStop",

  public = list(
    initialize = function(gs_elm) {
      private$.gs <- gs_elm
    }
  ),

  active = list(
    # Position 0.0–1.0
    position = function(value) {
      if (!missing(value)) {
        private$.gs$position <- value
        return(invisible(value))
      }
      private$.gs$position
    },

    # ColorFormat for this stop's color.
    # No-op setter allows chaining: stop$color$rgb <- RGBColor(...)
    color = function(value) {
      if (!missing(value)) return(invisible(NULL))
      if (is.null(private$.color_cache)) {
        private$.color_cache <- ColorFormat$new(private$.gs)
      }
      private$.color_cache
    }
  ),

  private = list(.gs = NULL, .color_cache = NULL)
)


# ============================================================================
# GradientStops — sequence of GradientStop objects
# ============================================================================

#' Gradient stops collection
#' @noRd
GradientStops <- R6::R6Class(
  "GradientStops",

  public = list(
    initialize = function(gsLst_elm) {
      private$.gsLst <- gsLst_elm
    },

    get = function(idx) {
      gs_list <- private$.gsLst$gs_lst
      if (idx < 1L || idx > length(gs_list)) {
        stop("gradient stop index out of range", call. = FALSE)
      }
      GradientStop$new(gs_list[[idx]])
    },

    to_list = function() {
      lapply(private$.gsLst$gs_lst, GradientStop$new)
    }
  ),

  private = list(.gsLst = NULL)
)

#' @export
length.GradientStops <- function(x) {
  length(x$.__enclos_env__$private$.gsLst$gs_lst)
}

#' @export
`[[.GradientStops` <- function(x, i) x$get(i)

#' @export
`[[<-.GradientStops` <- function(x, i, value) x


# ============================================================================
# FillFormat — access to fill properties on a shape or line element
# ============================================================================

#' Fill format accessor
#'
#' Wraps the shape-properties element (`<p:spPr>`) and exposes fill properties.
#' Access via `shape$fill`.
#'
#' @noRd
FillFormat <- R6::R6Class(
  "FillFormat",

  public = list(
    #' @param spPr A CT_ShapeProperties or similar element that contains fill.
    initialize = function(spPr) {
      private$.spPr <- spPr
    },

    # Apply a solid fill; call fore_color$rgb<- or fore_color$theme_color<- to set color.
    solid = function() {
      .remove_fill_elms(private$.spPr)
      a <- .nsmap[["a"]]
      fill_elm <- .insert_fill_child(private$.spPr,
                                     sprintf('<a:solidFill xmlns:a="%s"/>', a))
      invisible(self)
    },

    # Remove fill (transparent / inherit)
    background = function() {
      .remove_fill_elms(private$.spPr)
      invisible(self)
    },

    # Apply a gradient fill with default 2-stop linear gradient
    gradient = function() {
      .remove_fill_elms(private$.spPr)
      .insert_fill_child(private$.spPr, .default_gradFill_xml())
      invisible(self)
    },

    # Apply a pattern fill with the given preset
    patterned = function() {
      .remove_fill_elms(private$.spPr)
      a <- .nsmap[["a"]]
      .insert_fill_child(private$.spPr,
                         sprintf('<a:pattFill xmlns:a="%s" prst="pct5"/>', a))
      invisible(self)
    }
  ),

  active = list(
    # MSO_FILL type string for the current fill, or NULL
    type = function(value) {
      if (!missing(value)) return(invisible(NULL))
      elm <- .get_fill_elm(private$.spPr)
      if (is.null(elm)) return(NULL)
      if (inherits(elm, "CT_NoFill"))                   return(MSO_FILL$BACKGROUND)
      if (inherits(elm, "CT_SolidColorFillProperties")) return(MSO_FILL$SOLID)
      if (inherits(elm, "CT_GradientFillProperties"))   return(MSO_FILL$GRADIENT)
      if (inherits(elm, "CT_PatternFillProperties"))     return(MSO_FILL$PATTERNED)
      if (inherits(elm, "CT_BlipFillProperties"))        return(MSO_FILL$PICTURE)
      if (inherits(elm, "CT_GroupFillProperties"))       return(MSO_FILL$GROUP)
      NULL
    },

    # ColorFormat for the foreground color.
    # For solid fills: wraps the solidFill element itself.
    # For pattern fills: wraps the fgClr child, creating it if absent.
    # No-op setter allows chaining: shape$fill$fore_color$rgb <- RGBColor(...)
    fore_color = function(value) {
      if (!missing(value)) return(invisible(NULL))
      fill_elm <- .get_fill_elm(private$.spPr)
      if (is.null(fill_elm)) {
        stop("fore_color not available: no fill; call fill$solid() or fill$patterned() first",
             call. = FALSE)
      }
      if (inherits(fill_elm, "CT_SolidColorFillProperties")) {
        return(ColorFormat$new(fill_elm))
      }
      if (inherits(fill_elm, "CT_PatternFillProperties")) {
        return(ColorFormat$new(fill_elm$get_or_add_fgClr()))
      }
      stop("fore_color only available on solid or pattern fills", call. = FALSE)
    },

    # ColorFormat for the background (pattern fill) color.
    # For pattern fills: wraps the bgClr child, creating it if absent.
    # No-op setter allows chaining: shape$fill$back_color$rgb <- RGBColor(...)
    back_color = function(value) {
      if (!missing(value)) return(invisible(NULL))
      fill_elm <- .get_fill_elm(private$.spPr)
      if (is.null(fill_elm) || !inherits(fill_elm, "CT_PatternFillProperties")) {
        stop("back_color only available on pattern fills; call fill$patterned() first",
             call. = FALSE)
      }
      ColorFormat$new(fill_elm$get_or_add_bgClr())
    },

    # Pattern fill type string (prst attribute), or NULL if not a pattern fill
    pattern = function(value) {
      if (!missing(value)) {
        fill_elm <- .get_fill_elm(private$.spPr)
        if (is.null(fill_elm) || !inherits(fill_elm, "CT_PatternFillProperties")) {
          stop("pattern only available on pattern fills; call fill$patterned() first",
               call. = FALSE)
        }
        fill_elm$prst <- value
        return(invisible(value))
      }
      fill_elm <- .get_fill_elm(private$.spPr)
      if (is.null(fill_elm) || !inherits(fill_elm, "CT_PatternFillProperties")) {
        return(NULL)
      }
      fill_elm$prst
    },

    # GradientStops for a gradient fill, or NULL
    gradient_stops = function(value) {
      if (!missing(value)) return(invisible(NULL))
      fill_elm <- .get_fill_elm(private$.spPr)
      if (is.null(fill_elm) || !inherits(fill_elm, "CT_GradientFillProperties")) {
        return(NULL)
      }
      gsLst <- fill_elm$gsLst
      if (is.null(gsLst)) return(NULL)
      GradientStops$new(gsLst)
    },

    # Gradient angle in degrees (read/write), or NULL if not a gradient fill
    gradient_angle = function(value) {
      if (!missing(value)) {
        fill_elm <- .get_fill_elm(private$.spPr)
        if (is.null(fill_elm) || !inherits(fill_elm, "CT_GradientFillProperties")) {
          stop("gradient_angle only available on gradient fills", call. = FALSE)
        }
        lin <- fill_elm$get_or_add_lin()
        lin$ang <- value
        return(invisible(value))
      }
      fill_elm <- .get_fill_elm(private$.spPr)
      if (is.null(fill_elm) || !inherits(fill_elm, "CT_GradientFillProperties")) {
        return(NULL)
      }
      lin <- fill_elm$lin
      if (is.null(lin)) return(NULL)
      lin$ang
    }
  ),

  private = list(.spPr = NULL)
)

# Class-level factory: FillFormat$from_fill_parent(spPr) — mirrors Python classmethod.
FillFormat$from_fill_parent <- function(spPr) FillFormat$new(spPr)
