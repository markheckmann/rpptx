# DML color domain objects.
#
# Ported from python-pptx/src/pptx/dml/color.py.
# Provides RGBColor S3 class and ColorFormat R6 class.

# ============================================================================
# RGBColor — S3 class for an RGB color value
# ============================================================================

#' Create an RGB color value
#'
#' Represents a color as (red, green, blue) components in 0–255 range.
#'
#' @param r Integer 0–255 red component.
#' @param g Integer 0–255 green component.
#' @param b Integer 0–255 blue component.
#' @return An `RGBColor` object.
#'
#' @export
RGBColor <- function(r, g, b) {
  r <- as.integer(r); g <- as.integer(g); b <- as.integer(b)
  if (any(c(r, g, b) < 0L | c(r, g, b) > 255L)) {
    stop("RGB components must be in range 0-255", call. = FALSE)
  }
  structure(list(r = r, g = g, b = b), class = "RGBColor")
}

#' @export
print.RGBColor <- function(x, ...) {
  cat(sprintf("RGBColor(%d, %d, %d) [#%s]\n", x$r, x$g, x$b, as.character(x)))
  invisible(x)
}

#' @export
as.character.RGBColor <- function(x, ...) {
  sprintf("%02X%02X%02X", x$r, x$g, x$b)
}

#' @export
`==.RGBColor` <- function(a, b) {
  if (!inherits(b, "RGBColor")) return(FALSE)
  a$r == b$r && a$g == b$g && a$b == b$b
}

#' Create RGBColor from a 6-character hex string
#'
#' @param hex_str A 6-character hex string like "FF0000" or "#FF0000".
#' @return An `RGBColor` object.
#' @export
RGBColor_from_str <- function(hex_str) {
  hex_str <- toupper(sub("^#", "", hex_str))
  if (nchar(hex_str) != 6L) stop("hex_str must be 6 hex digits", call. = FALSE)
  r <- strtoi(substr(hex_str, 1, 2), 16L)
  g <- strtoi(substr(hex_str, 3, 4), 16L)
  b <- strtoi(substr(hex_str, 5, 6), 16L)
  RGBColor(r, g, b)
}


# ============================================================================
# ColorFormat — access to color properties on a shape fill/line element
# ============================================================================

#' Color format accessor
#'
#' Wraps the "color choice parent" element (e.g. `<a:solidFill>`, `<a:fgClr>`,
#' `<a:bgClr>`, or a gradient stop `<a:gs>`) and exposes color properties.
#'
#' Access via `FillFormat$fore_color`, `FillFormat$back_color`, or
#' `LineFormat$color`.
#'
#' @noRd
ColorFormat <- R6::R6Class(
  "ColorFormat",

  public = list(
    #' @param xPr A ColorChoiceParent element (e.g. CT_SolidColorFillProperties).
    initialize = function(xPr) {
      private$.xPr <- xPr
    }
  ),

  active = list(
    # MSO_COLOR_TYPE member: "rgb", "scheme", or NULL if no color set
    type = function(value) {
      if (!missing(value)) return(invisible(NULL))
      choice <- private$.xPr$eg_colorChoice
      if (is.null(choice)) return(NULL)
      if (inherits(choice, "CT_SRgbColor"))   return(MSO_COLOR_TYPE$RGB)
      if (inherits(choice, "CT_SchemeColor"))  return(MSO_COLOR_TYPE$SCHEME)
      if (inherits(choice, "CT_HslColor"))     return(MSO_COLOR_TYPE$HSL)
      if (inherits(choice, "CT_ScRgbColor"))   return(MSO_COLOR_TYPE$SCRGB)
      if (inherits(choice, "CT_PresetColor"))  return(MSO_COLOR_TYPE$PRESET)
      if (inherits(choice, "CT_SystemColor"))  return(MSO_COLOR_TYPE$SYSTEM)
      NULL
    },

    # RGBColor for this color (getter); set to RGBColor to define solid RGB color.
    # Returns NULL if color is not an sRGB color.
    rgb = function(value) {
      if (!missing(value)) {
        hex_str <- as.character(value)
        private$.xPr$get_or_add_srgbClr(hex_str)
        return(invisible(value))
      }
      srgb <- private$.xPr$srgbClr
      if (is.null(srgb)) {
        # Fallback: check sysClr lastClr
        choice <- private$.xPr$eg_colorChoice
        if (inherits(choice, "CT_SystemColor")) {
          lc <- choice$lastClr
          if (!is.null(lc)) return(RGBColor_from_str(lc))
        }
        return(NULL)
      }
      RGBColor_from_str(srgb$val)
    },

    # Theme color string (MSO_THEME_COLOR value) for this color.
    # Set to an MSO_THEME_COLOR value string to apply a theme color.
    theme_color = function(value) {
      if (!missing(value)) {
        theme_val <- if (is.null(value)) NULL else as.character(value)
        if (is.null(theme_val)) {
          private$.xPr$clear_color_choice()
        } else {
          private$.xPr$get_or_add_schemeClr(theme_val)
        }
        return(invisible(value))
      }
      scheme <- private$.xPr$schemeClr
      if (is.null(scheme)) return(NULL)
      scheme$val
    },

    # Tint: fraction by which the color is lightened toward white (0.0–1.0).
    # NULL when no tint is set; only valid on theme (scheme) colors.
    tint = function(value) {
      if (!missing(value)) {
        scheme <- private$.xPr$schemeClr
        if (is.null(scheme)) stop("tint only applies to scheme (theme) colors", call. = FALSE)
        scheme[["_remove_lumMod"]]()
        scheme[["_remove_lumOff"]]()
        if (!is.null(value) && value > 0.0) {
          lm <- scheme$get_or_add_lumMod(); lm$val <- 1.0 - value
          lo <- scheme$get_or_add_lumOff(); lo$val <- value
        }
        return(invisible(value))
      }
      scheme <- private$.xPr$schemeClr
      if (is.null(scheme)) return(NULL)
      lumOff <- scheme$lumOff
      if (!is.null(lumOff)) return(lumOff$val)
      NULL
    },

    # Shade: fraction by which the color is darkened toward black (0.0–1.0).
    # NULL when no shade is set; only valid on theme (scheme) colors.
    shade = function(value) {
      if (!missing(value)) {
        scheme <- private$.xPr$schemeClr
        if (is.null(scheme)) stop("shade only applies to scheme (theme) colors", call. = FALSE)
        scheme[["_remove_lumMod"]]()
        scheme[["_remove_lumOff"]]()
        if (!is.null(value) && value > 0.0) {
          lm <- scheme$get_or_add_lumMod(); lm$val <- 1.0 - value
        }
        return(invisible(value))
      }
      scheme <- private$.xPr$schemeClr
      if (is.null(scheme)) return(NULL)
      lumMod <- scheme$lumMod
      lumOff <- scheme$lumOff
      if (!is.null(lumMod) && is.null(lumOff)) return(1.0 - lumMod$val)
      NULL
    },

    # Brightness adjustment as a float in -1.0..1.0 range.
    # Positive values add lightness (lumOff), negative values reduce (lumMod).
    # NULL means no brightness adjustment or not a scheme color.
    brightness = function(value) {
      if (!missing(value)) {
        scheme <- private$.xPr$schemeClr
        if (is.null(scheme)) {
          stop("brightness can only be set on theme (scheme) colors", call. = FALSE)
        }
        # Remove existing lumMod / lumOff
        scheme[["_remove_lumMod"]]()
        scheme[["_remove_lumOff"]]()
        if (!is.null(value) && value != 0.0) {
          if (value > 0) {
            # Brighten: add lumMod=1-value, lumOff=value
            lm <- scheme$get_or_add_lumMod()
            lm$val <- 1.0 - value
            lo <- scheme$get_or_add_lumOff()
            lo$val <- value
          } else {
            # Darken: add lumMod=1+value (value is negative)
            lm <- scheme$get_or_add_lumMod()
            lm$val <- 1.0 + value
          }
        }
        return(invisible(value))
      }
      scheme <- private$.xPr$schemeClr
      if (is.null(scheme)) return(NULL)
      lumOff <- scheme$lumOff
      if (!is.null(lumOff)) return(lumOff$val)
      lumMod <- scheme$lumMod
      if (!is.null(lumMod)) return(lumMod$val - 1.0)
      0.0
    }
  ),

  private = list(.xPr = NULL)
)
