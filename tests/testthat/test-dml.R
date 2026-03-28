# Tests for DML — Colors, Fills, Lines.
# Covers: RGBColor, ColorFormat, FillFormat, LineFormat,
#         oxml-dml-color, oxml-dml-fill, oxml-dml-line, enum-dml.

pptx_path <- function(name) system.file("test_files", name, package = "rpptx")

# Helper: a fresh Shape with an empty <p:spPr/>
new_shape_with_empty_spPr <- function() {
  prs    <- pptx_presentation(pptx_path("no-slides.pptx"))
  layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
  slide  <- prs$slides$add_slide(layout)
  slide$shapes$add_shape(
    MSO_AUTO_SHAPE_TYPE$RECTANGLE,
    Inches(1), Inches(1), Inches(2), Inches(2)
  )
}


# ============================================================================
# Enum: MSO_FILL
# ============================================================================

describe("MSO_FILL", {
  it("has SOLID", { expect_equal(MSO_FILL$SOLID, "solid") })
  it("has GRADIENT", { expect_equal(MSO_FILL$GRADIENT, "gradient") })
  it("has BACKGROUND", { expect_equal(MSO_FILL$BACKGROUND, "background") })
  it("has PATTERNED", { expect_equal(MSO_FILL$PATTERNED, "patterned") })
  it("has PICTURE", { expect_equal(MSO_FILL$PICTURE, "picture") })
  it("has GROUP", { expect_equal(MSO_FILL$GROUP, "group") })
})


# ============================================================================
# Enum: MSO_LINE_DASH_STYLE
# ============================================================================

describe("MSO_LINE_DASH_STYLE", {
  it("has SOLID", { expect_equal(MSO_LINE_DASH_STYLE$SOLID, "solid") })
  it("has DASH", { expect_equal(MSO_LINE_DASH_STYLE$DASH, "dash") })
  it("has ROUND_DOT", { expect_equal(MSO_LINE_DASH_STYLE$ROUND_DOT, "sysDot") })
})


# ============================================================================
# Enum: MSO_THEME_COLOR
# ============================================================================

describe("MSO_THEME_COLOR", {
  it("has ACCENT_1", { expect_equal(MSO_THEME_COLOR$ACCENT_1, "accent1") })
  it("has DARK_1",   { expect_equal(MSO_THEME_COLOR$DARK_1,   "dk1") })
  it("has NOT_THEME_COLOR as NULL", { expect_null(MSO_THEME_COLOR$NOT_THEME_COLOR) })
})


# ============================================================================
# RGBColor
# ============================================================================

describe("RGBColor()", {
  it("creates an RGBColor", {
    col <- RGBColor(255L, 0L, 128L)
    expect_s3_class(col, "RGBColor")
  })

  it("stores r, g, b", {
    col <- RGBColor(10L, 20L, 30L)
    expect_equal(col$r, 10L)
    expect_equal(col$g, 20L)
    expect_equal(col$b, 30L)
  })

  it("as.character() returns 6-char uppercase hex", {
    expect_equal(as.character(RGBColor(255L, 0L, 0L)),  "FF0000")
    expect_equal(as.character(RGBColor(0L, 255L, 0L)),  "00FF00")
    expect_equal(as.character(RGBColor(0L, 0L, 255L)),  "0000FF")
    expect_equal(as.character(RGBColor(16L, 32L, 48L)), "102030")
  })

  it("errors on out-of-range values", {
    expect_error(RGBColor(256L, 0L, 0L))
    expect_error(RGBColor(-1L, 0L, 0L))
  })

  it("equality works", {
    expect_true(RGBColor(1L, 2L, 3L) == RGBColor(1L, 2L, 3L))
    expect_false(RGBColor(1L, 2L, 3L) == RGBColor(1L, 2L, 4L))
  })
})

describe("RGBColor_from_str()", {
  it("parses 6-char hex string", {
    col <- RGBColor_from_str("FF0000")
    expect_equal(col$r, 255L)
    expect_equal(col$g, 0L)
    expect_equal(col$b, 0L)
  })

  it("parses with leading #", {
    col <- RGBColor_from_str("#00FF80")
    expect_equal(col$g, 255L)
    expect_equal(col$b, 128L)
  })

  it("is case-insensitive", {
    col <- RGBColor_from_str("ff8800")
    expect_equal(col$r, 255L)
    expect_equal(col$g, 136L)
  })

  it("round-trips with as.character", {
    original <- RGBColor(100L, 150L, 200L)
    recovered <- RGBColor_from_str(as.character(original))
    expect_true(original == recovered)
  })
})


# ============================================================================
# CT_SRgbColor / CT_SchemeColor elements
# ============================================================================

describe("CT_SRgbColor element", {
  it("can be created by parsing XML", {
    xml_str <- sprintf('<a:srgbClr xmlns:a="%s" val="FF0000"/>', .nsmap[["a"]])
    elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    expect_true(inherits(elm, "CT_SRgbColor"))
    expect_equal(elm$val, "FF0000")
  })

  it("val is settable", {
    xml_str <- sprintf('<a:srgbClr xmlns:a="%s" val="000000"/>', .nsmap[["a"]])
    elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    elm$val <- "AABBCC"
    expect_equal(elm$val, "AABBCC")
  })
})

describe("CT_SchemeColor element", {
  it("can be created with val attribute", {
    xml_str <- sprintf('<a:schemeClr xmlns:a="%s" val="accent1"/>', .nsmap[["a"]])
    elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    expect_true(inherits(elm, "CT_SchemeColor"))
    expect_equal(elm$val, "accent1")
  })

  it("get_or_add_lumMod creates lumMod child", {
    xml_str <- sprintf('<a:schemeClr xmlns:a="%s" val="accent1"/>', .nsmap[["a"]])
    elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    lm <- elm$get_or_add_lumMod()
    expect_true(inherits(lm, "CT_Percentage"))
  })

  it("lumOff starts NULL", {
    xml_str <- sprintf('<a:schemeClr xmlns:a="%s" val="accent1"/>', .nsmap[["a"]])
    elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    expect_null(elm$lumOff)
  })
})


# ============================================================================
# CT_LineProperties element
# ============================================================================

describe("CT_LineProperties element", {
  it("is registered as CT_LineProperties", {
    xml_str <- sprintf('<a:ln xmlns:a="%s"/>', .nsmap[["a"]])
    elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    expect_true(inherits(elm, "CT_LineProperties"))
  })

  it("w attribute round-trips as Emu", {
    xml_str <- sprintf('<a:ln xmlns:a="%s" w="12700"/>', .nsmap[["a"]])
    elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    expect_equal(as.integer(elm$w), 12700L)
  })

  it("get_or_add_prstDash creates child", {
    xml_str <- sprintf('<a:ln xmlns:a="%s"/>', .nsmap[["a"]])
    elm <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    pd <- elm$get_or_add_prstDash()
    expect_true(inherits(pd, "CT_PresetLineDashProperties"))
  })
})


# ============================================================================
# ColorFormat
# ============================================================================

describe("ColorFormat — RGB color", {
  make_solidFill <- function() {
    xml_str <- sprintf('<a:solidFill xmlns:a="%s"/>', .nsmap[["a"]])
    wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
  }

  it("type is NULL when no color set", {
    sf <- make_solidFill()
    cf <- ColorFormat$new(sf)
    expect_null(cf$type)
  })

  it("setting rgb creates srgbClr child", {
    sf <- make_solidFill()
    cf <- ColorFormat$new(sf)
    cf$rgb <- RGBColor(255L, 0L, 0L)
    expect_equal(cf$type, MSO_COLOR_TYPE$RGB)
    expect_true(cf$rgb == RGBColor(255L, 0L, 0L))
  })

  it("rgb getter returns NULL when schemeClr set", {
    sf <- make_solidFill()
    cf <- ColorFormat$new(sf)
    cf$theme_color <- MSO_THEME_COLOR$ACCENT_1
    expect_null(cf$rgb)
  })
})

describe("ColorFormat — theme color", {
  make_solidFill <- function() {
    xml_str <- sprintf('<a:solidFill xmlns:a="%s"/>', .nsmap[["a"]])
    wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
  }

  it("type is SCHEME after setting theme_color", {
    sf <- make_solidFill()
    cf <- ColorFormat$new(sf)
    cf$theme_color <- MSO_THEME_COLOR$ACCENT_1
    expect_equal(cf$type, MSO_COLOR_TYPE$SCHEME)
  })

  it("theme_color getter returns the val string", {
    sf <- make_solidFill()
    cf <- ColorFormat$new(sf)
    cf$theme_color <- MSO_THEME_COLOR$DARK_1
    expect_equal(cf$theme_color, "dk1")
  })

  it("setting theme_color replaces existing srgbClr", {
    sf <- make_solidFill()
    cf <- ColorFormat$new(sf)
    cf$rgb <- RGBColor(255L, 0L, 0L)
    cf$theme_color <- MSO_THEME_COLOR$ACCENT_2
    expect_equal(cf$type, MSO_COLOR_TYPE$SCHEME)
    expect_null(cf$rgb)
  })
})

describe("ColorFormat — brightness", {
  make_scheme_solidFill <- function(val = "accent1") {
    a <- .nsmap[["a"]]
    xml_str <- sprintf('<a:solidFill xmlns:a="%s"><a:schemeClr val="%s"/></a:solidFill>',
                       a, val)
    wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
  }

  it("brightness 0 on plain schemeClr returns 0.0", {
    sf <- make_scheme_solidFill()
    cf <- ColorFormat$new(sf)
    expect_equal(cf$brightness, 0.0)
  })

  it("setting positive brightness adds lumMod and lumOff", {
    sf <- make_scheme_solidFill()
    cf <- ColorFormat$new(sf)
    cf$brightness <- 0.25
    expect_equal(cf$brightness, 0.25, tolerance = 1e-5)
  })

  it("setting negative brightness adds only lumMod", {
    sf <- make_scheme_solidFill()
    cf <- ColorFormat$new(sf)
    cf$brightness <- -0.4
    expect_equal(cf$brightness, -0.4, tolerance = 1e-5)
  })

  it("errors when called on non-scheme color", {
    a <- .nsmap[["a"]]
    xml_str <- sprintf('<a:solidFill xmlns:a="%s"><a:srgbClr val="FF0000"/></a:solidFill>', a)
    sf <- wrap_element(xml2::xml_root(xml2::read_xml(xml_str)))
    cf <- ColorFormat$new(sf)
    expect_error(cf$brightness <- 0.5, "brightness")
  })
})


# ============================================================================
# FillFormat — via shape$fill
# ============================================================================

describe("FillFormat — type", {
  it("type is NULL for no explicit fill on a fresh shape", {
    shape <- new_shape_with_empty_spPr()
    # autoshape XML has no fill element, so type should be NULL
    expect_null(shape$fill$type)
  })
})

describe("FillFormat$solid()", {
  it("sets fill type to SOLID", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$solid()
    expect_equal(shape$fill$type, MSO_FILL$SOLID)
  })

  it("fore_color is accessible after solid()", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$solid()
    cf <- shape$fill$fore_color
    expect_true(inherits(cf, "ColorFormat"))
  })

  it("fore_color$rgb can be set", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$solid()
    shape$fill$fore_color$rgb <- RGBColor(0L, 128L, 255L)
    expect_true(shape$fill$fore_color$rgb == RGBColor(0L, 128L, 255L))
  })

  it("setting type via no-op setter is a no-op", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$solid()
    shape$fill$type <- "anything"  # should be ignored
    expect_equal(shape$fill$type, MSO_FILL$SOLID)
  })
})

describe("FillFormat$background()", {
  it("removes fill (type becomes NULL)", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$solid()
    shape$fill$background()
    expect_null(shape$fill$type)
  })
})

describe("FillFormat$gradient()", {
  it("sets fill type to GRADIENT", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$gradient()
    expect_equal(shape$fill$type, MSO_FILL$GRADIENT)
  })

  it("gradient_stops returns GradientStops", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$gradient()
    gs <- shape$fill$gradient_stops
    expect_true(inherits(gs, "GradientStops"))
  })

  it("default gradient has 2 stops", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$gradient()
    expect_equal(length(shape$fill$gradient_stops), 2L)
  })

  it("gradient_stops[[1]] is a GradientStop", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$gradient()
    stop1 <- shape$fill$gradient_stops[[1]]
    expect_true(inherits(stop1, "GradientStop"))
  })

  it("gradient_stops positions are 0.0 and 1.0", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$gradient()
    stops <- shape$fill$gradient_stops
    expect_equal(stops[[1]]$position, 0.0, tolerance = 1e-5)
    expect_equal(stops[[2]]$position, 1.0, tolerance = 1e-5)
  })
})

describe("FillFormat$patterned()", {
  it("sets fill type to PATTERNED", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$patterned()
    expect_equal(shape$fill$type, MSO_FILL$PATTERNED)
  })

  it("pattern attribute is readable", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$patterned()
    expect_false(is.null(shape$fill$pattern))
  })

  it("pattern attribute is settable", {
    shape <- new_shape_with_empty_spPr()
    shape$fill$patterned()
    shape$fill$pattern <- MSO_PATTERN_TYPE$CROSS
    expect_equal(shape$fill$pattern, "cross")
  })
})

describe("FillFormat — fore_color errors on non-solid", {
  it("errors when fill is not solid", {
    shape <- new_shape_with_empty_spPr()
    expect_error(shape$fill$fore_color, "solid")
  })
})


# ============================================================================
# LineFormat — via shape$line
# ============================================================================

describe("LineFormat — width", {
  it("width is NULL when no a:ln element", {
    shape <- new_shape_with_empty_spPr()
    expect_null(shape$line$width)
  })

  it("width can be set and retrieved", {
    shape <- new_shape_with_empty_spPr()
    shape$line$width <- Pt(1)
    expect_equal(as.integer(shape$line$width), as.integer(Pt(1)))
  })

  it("setting width creates a:ln element", {
    shape <- new_shape_with_empty_spPr()
    shape$line$width <- Pt(2)
    ln <- shape$element$ln
    expect_false(is.null(ln))
  })
})

describe("LineFormat — dash_style", {
  it("dash_style is NULL with no line element", {
    shape <- new_shape_with_empty_spPr()
    expect_null(shape$line$dash_style)
  })

  it("dash_style can be set and retrieved", {
    shape <- new_shape_with_empty_spPr()
    shape$line$dash_style <- MSO_LINE_DASH_STYLE$DASH
    expect_equal(shape$line$dash_style, "dash")
  })
})

describe("LineFormat — color", {
  it("color returns a ColorFormat", {
    shape <- new_shape_with_empty_spPr()
    cf <- shape$line$color
    expect_true(inherits(cf, "ColorFormat"))
  })

  it("line color rgb can be set", {
    shape <- new_shape_with_empty_spPr()
    shape$line$color$rgb <- RGBColor(255L, 0L, 0L)
    expect_true(shape$line$color$rgb == RGBColor(255L, 0L, 0L))
  })

  it("line fill type is SOLID after setting color", {
    shape <- new_shape_with_empty_spPr()
    shape$line$color$rgb <- RGBColor(0L, 255L, 0L)
    expect_equal(shape$line$fill, MSO_FILL$SOLID)
  })

  it("no-op setter on line is a no-op", {
    shape <- new_shape_with_empty_spPr()
    shape$line$color$rgb <- RGBColor(0L, 0L, 255L)
    shape$line <- "ignored"  # no-op setter
    expect_true(shape$line$color$rgb == RGBColor(0L, 0L, 255L))
  })
})
