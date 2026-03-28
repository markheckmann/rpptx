# Simple-type value converters for XML attributes.
#
# Ported from python-pptx/src/pptx/oxml/simpletypes.py. Each simple type
# is a list with `from_xml` and `to_xml` functions that convert between
# XML string values and R values.

# --- Helper to create a simple type ---

#' Create a simple type converter
#' @param from_xml Function(str) -> R value.
#' @param to_xml Function(R value) -> str.
#' @return A list with `from_xml` and `to_xml` elements.
#' @keywords internal
simple_type <- function(from_xml, to_xml) {
  list(from_xml = from_xml, to_xml = to_xml)
}


# ============================================================================
# XSD base types
# ============================================================================

#' @keywords internal
XsdString <- simple_type(
  from_xml = function(x) x,
  to_xml = function(x) as.character(x)
)

#' @keywords internal
XsdBoolean <- simple_type(
  from_xml = function(x) {
    if (!(x %in% c("1", "0", "true", "false"))) {
      invalid_xml_error(sprintf(
        "value must be one of '1', '0', 'true' or 'false', got '%s'", x
      ))
    }
    x %in% c("1", "true")
  },
  to_xml = function(x) if (isTRUE(x)) "1" else "0"
)

#' @keywords internal
XsdInt <- simple_type(
  from_xml = function(x) as.integer(x),
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
XsdUnsignedInt <- simple_type(
  from_xml = function(x) as.integer(x),
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
XsdUnsignedByte <- simple_type(
  from_xml = function(x) as.integer(x),
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
XsdUnsignedShort <- simple_type(
  from_xml = function(x) as.integer(x),
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
XsdLong <- simple_type(
  from_xml = function(x) as.numeric(x),
  to_xml = function(x) format(x, scientific = FALSE)
)

#' @keywords internal
XsdDouble <- simple_type(
  from_xml = function(x) as.numeric(x),
  to_xml = function(x) as.character(as.numeric(x))
)

#' @keywords internal
XsdAnyUri <- XsdString

#' @keywords internal
XsdId <- XsdString

#' @keywords internal
XsdToken <- XsdString


# ============================================================================
# ST_* types used across the library
# ============================================================================

#' @keywords internal
ST_RelationshipId <- XsdString

#' @keywords internal
ST_ContentType <- XsdString

#' @keywords internal
ST_Extension <- XsdString

#' @keywords internal
ST_TargetMode <- simple_type(
  from_xml = function(x) x,
  to_xml = function(x) {
    if (!(x %in% c("External", "Internal"))) {
      stop(sprintf("must be 'Internal' or 'External', got '%s'", x), call. = FALSE)
    }
    x
  }
)

#' @keywords internal
ST_Coordinate <- simple_type(
  from_xml = function(x) {
    if (grepl("[impc]", x)) {
      return(ST_UniversalMeasure$from_xml(x))
    }
    Emu(as.integer(x))
  },
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_Coordinate32 <- simple_type(
  from_xml = function(x) {
    if (grepl("[impc]", x)) {
      return(ST_UniversalMeasure$from_xml(x))
    }
    Emu(as.integer(x))
  },
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_PositiveCoordinate <- simple_type(
  from_xml = function(x) Emu(as.integer(x)),
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_SlideSizeCoordinate <- simple_type(
  from_xml = function(x) Emu(as.integer(x)),
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_SlideId <- XsdUnsignedInt

#' @keywords internal
ST_DrawingElementId <- XsdUnsignedInt

#' @keywords internal
ST_Angle <- simple_type(
  from_xml = function(x) {
    rot <- as.integer(x) %% (360L * 60000L)
    rot / 60000.0
  },
  to_xml = function(x) {
    rot <- as.integer(round(x * 60000)) %% (360L * 60000L)
    as.character(rot)
  }
)

#' @keywords internal
ST_LineWidth <- simple_type(
  from_xml = function(x) Emu(as.integer(x)),
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_HexColorRGB <- simple_type(
  from_xml = function(x) toupper(x),
  to_xml = function(x) toupper(x)
)

#' @keywords internal
ST_Percentage <- simple_type(
  from_xml = function(x) {
    if (grepl("%", x, fixed = TRUE)) {
      return(as.numeric(sub("%", "", x)) / 100.0)
    }
    as.integer(x) / 100000.0
  },
  to_xml = function(x) as.character(as.integer(round(x * 100000.0)))
)

#' @keywords internal
ST_PositiveFixedPercentage <- simple_type(
  from_xml = function(x) ST_Percentage$from_xml(x),
  to_xml = function(x) ST_Percentage$to_xml(x)
)

#' @keywords internal
ST_TextFontSize <- simple_type(
  from_xml = function(x) as.integer(x),
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_TextIndentLevelType <- simple_type(
  from_xml = function(x) as.integer(x),
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_TextTypeface <- XsdString

#' @keywords internal
ST_TextWrappingType <- simple_type(
  from_xml = function(x) x,
  to_xml = function(x) x
)

#' @keywords internal
ST_TextFontScalePercentOrPercentString <- simple_type(
  from_xml = function(x) {
    if (endsWith(x, "%")) return(as.numeric(sub("%", "", x)))
    as.integer(x) / 1000.0
  },
  to_xml = function(x) as.character(as.integer(x * 1000.0))
)

#' @keywords internal
ST_TextSpacingPercentOrPercentString <- simple_type(
  from_xml = function(x) {
    if (endsWith(x, "%")) {
      return(as.numeric(sub("%", "", x)) / 100.0)
    }
    as.integer(x) / 100000.0
  },
  to_xml = function(x) as.character(as.integer(round(x * 100000.0)))
)

#' @keywords internal
ST_TextSpacingPoint <- simple_type(
  from_xml = function(x) Centipoints(as.integer(x)),
  to_xml = function(x) as.character(as_centipoints(x))
)

#' @keywords internal
ST_PlaceholderSize <- simple_type(
  from_xml = function(x) x,
  to_xml = function(x) x
)

#' @keywords internal
ST_Direction <- simple_type(
  from_xml = function(x) x,
  to_xml = function(x) x
)

#' @keywords internal
ST_UniversalMeasure <- simple_type(
  from_xml = function(x) {
    float_part <- substr(x, 1, nchar(x) - 2)
    units_part <- substr(x, nchar(x) - 1, nchar(x))
    quantity <- as.numeric(float_part)
    multiplier <- switch(units_part,
      "mm" = 36000, "cm" = 360000, "in" = 914400,
      "pt" = 12700, "pc" = 152400, "pi" = 152400,
      stop(sprintf("Unknown unit: '%s'", units_part), call. = FALSE)
    )
    Emu(as.integer(round(quantity * multiplier)))
  },
  to_xml = function(x) as.character(as.integer(x))
)

# Chart-related simple types

#' @keywords internal
ST_BarDir <- simple_type(from_xml = function(x) x, to_xml = function(x) x)

#' @keywords internal
ST_Grouping <- simple_type(from_xml = function(x) x, to_xml = function(x) x)

#' @keywords internal
ST_Orientation <- simple_type(from_xml = function(x) x, to_xml = function(x) x)

#' @keywords internal
ST_AxisUnit <- simple_type(
  from_xml = function(x) as.numeric(x),
  to_xml = function(x) as.character(as.numeric(x))
)

#' @keywords internal
ST_BubbleScale <- simple_type(
  from_xml = function(x) {
    if (grepl("%", x, fixed = TRUE)) return(as.integer(sub("%", "", x)))
    as.integer(x)
  },
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_GapAmount <- ST_BubbleScale

#' @keywords internal
ST_Overlap <- simple_type(
  from_xml = function(x) {
    if (grepl("%", x, fixed = TRUE)) return(as.integer(sub("%", "", x)))
    as.integer(x)
  },
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_LblOffset <- simple_type(
  from_xml = function(x) {
    if (endsWith(x, "%")) return(as.integer(sub("%", "", x)))
    as.integer(x)
  },
  to_xml = function(x) as.character(as.integer(x))
)

#' @keywords internal
ST_LayoutMode <- simple_type(from_xml = function(x) x, to_xml = function(x) x)

#' @keywords internal
ST_MarkerSize <- XsdUnsignedByte

#' @keywords internal
ST_Style <- XsdUnsignedByte

# Simple pass-through types for chart XML string enums
# These read/write the XML string value directly. The enum list constants
# (XL_AXIS_CROSSES, XL_LEGEND_POSITION, etc.) use the same string values.

#' @keywords internal
ST_AxisCrosses <- simple_type(from_xml = function(x) x, to_xml = function(x) as.character(x))

#' @keywords internal
ST_DataLabelPosition <- simple_type(from_xml = function(x) x, to_xml = function(x) as.character(x))

#' @keywords internal
ST_LegendPosition <- simple_type(from_xml = function(x) x, to_xml = function(x) as.character(x))

#' @keywords internal
ST_MarkerStyle <- simple_type(from_xml = function(x) x, to_xml = function(x) as.character(x))

#' @keywords internal
ST_TickLabelPosition <- simple_type(from_xml = function(x) x, to_xml = function(x) as.character(x))

#' @keywords internal
ST_TickMark <- simple_type(from_xml = function(x) x, to_xml = function(x) as.character(x))

#' @keywords internal
ST_PositiveFixedAngle <- simple_type(
  from_xml = function(x) ST_Angle$from_xml(x),
  to_xml = function(x) {
    degrees <- x
    if (degrees < 0.0) {
      degrees <- (degrees %% -360) + 360
    } else if (degrees > 0.0) {
      degrees <- degrees %% 360
    }
    as.character(as.integer(round(degrees * 60000)))
  }
)
