# Tests for action.R, enum-action.R, enum-text.R
#
# Phase 10: ActionSetting, ShapeHyperlink, PP_ACTION_TYPE, text enums.

pptx_path <- function(name) system.file("test_files", name, package = "rpptx")

# ============================================================================
# Helper: build a CT_NonVisualDrawingProps element from XML string
# ============================================================================

make_cNvPr <- function(inner_xml = "") {
  ns_p <- .nsmap[["p"]]
  ns_a <- .nsmap[["a"]]
  ns_r <- .nsmap[["r"]]
  xml_str <- sprintf(
    '<p:cNvPr xmlns:p="%s" xmlns:a="%s" xmlns:r="%s" id="1" name="Shape 1">%s</p:cNvPr>',
    ns_p, ns_a, ns_r, inner_xml
  )
  rpptx_parse_xml(xml_str)
}

make_hlink_click_xml <- function(rId = "rId1", action = NULL) {
  action_attr <- if (!is.null(action)) {
    # Escape & → &amp; for XML validity in attribute values
    xml_action <- gsub("&", "&amp;", action, fixed = TRUE)
    sprintf(' action="%s"', xml_action)
  } else ""
  sprintf('<a:hlinkClick r:id="%s"%s/>', rId, action_attr)
}


# ============================================================================
# PP_ACTION_TYPE enum
# ============================================================================

describe("PP_ACTION_TYPE", {
  it("has NONE = 0L", {
    expect_identical(PP_ACTION_TYPE$NONE, 0L)
  })

  it("has HYPERLINK = 7L", {
    expect_identical(PP_ACTION_TYPE$HYPERLINK, 7L)
  })

  it("has FIRST_SLIDE = 3L", {
    expect_identical(PP_ACTION_TYPE$FIRST_SLIDE, 3L)
  })

  it("has LAST_SLIDE = 4L", {
    expect_identical(PP_ACTION_TYPE$LAST_SLIDE, 4L)
  })

  it("has NEXT_SLIDE = 1L", {
    expect_identical(PP_ACTION_TYPE$NEXT_SLIDE, 1L)
  })

  it("has PREVIOUS_SLIDE = 2L", {
    expect_identical(PP_ACTION_TYPE$PREVIOUS_SLIDE, 2L)
  })

  it("has NAMED_SLIDE = 101L", {
    expect_identical(PP_ACTION_TYPE$NAMED_SLIDE, 101L)
  })

  it("has END_SHOW = 6L", {
    expect_identical(PP_ACTION_TYPE$END_SHOW, 6L)
  })

  it("has RUN_MACRO = 8L", {
    expect_identical(PP_ACTION_TYPE$RUN_MACRO, 8L)
  })
})


# ============================================================================
# PP_PARAGRAPH_ALIGNMENT enum
# ============================================================================

describe("PP_PARAGRAPH_ALIGNMENT", {
  it("has CENTER = 'ctr'", {
    expect_equal(PP_PARAGRAPH_ALIGNMENT$CENTER, "ctr")
  })

  it("has LEFT = 'l'", {
    expect_equal(PP_PARAGRAPH_ALIGNMENT$LEFT, "l")
  })

  it("has RIGHT = 'r'", {
    expect_equal(PP_PARAGRAPH_ALIGNMENT$RIGHT, "r")
  })

  it("has JUSTIFY = 'just'", {
    expect_equal(PP_PARAGRAPH_ALIGNMENT$JUSTIFY, "just")
  })

  it("PP_ALIGN is an alias for PP_PARAGRAPH_ALIGNMENT", {
    expect_identical(PP_ALIGN, PP_PARAGRAPH_ALIGNMENT)
  })
})


# ============================================================================
# MSO_AUTO_SIZE enum
# ============================================================================

describe("MSO_AUTO_SIZE", {
  it("has NONE = 0L", {
    expect_identical(MSO_AUTO_SIZE$NONE, 0L)
  })

  it("has SHAPE_TO_FIT_TEXT = 1L", {
    expect_identical(MSO_AUTO_SIZE$SHAPE_TO_FIT_TEXT, 1L)
  })

  it("has TEXT_TO_FIT_SHAPE = 2L", {
    expect_identical(MSO_AUTO_SIZE$TEXT_TO_FIT_SHAPE, 2L)
  })

  it("has MIXED = -2L", {
    expect_identical(MSO_AUTO_SIZE$MIXED, -2L)
  })
})


# ============================================================================
# MSO_VERTICAL_ANCHOR enum
# ============================================================================

describe("MSO_VERTICAL_ANCHOR", {
  it("has TOP = 't'", {
    expect_equal(MSO_VERTICAL_ANCHOR$TOP, "t")
  })

  it("has MIDDLE = 'ctr'", {
    expect_equal(MSO_VERTICAL_ANCHOR$MIDDLE, "ctr")
  })

  it("has BOTTOM = 'b'", {
    expect_equal(MSO_VERTICAL_ANCHOR$BOTTOM, "b")
  })
})


# ============================================================================
# CT_Hyperlink — action_verb() and action_fields()
# ============================================================================

describe("CT_Hyperlink$action_verb()", {
  it("returns NULL when no action attribute", {
    cNvPr <- make_cNvPr(make_hlink_click_xml())
    hlink <- cNvPr$hlinkClick
    expect_null(hlink$action_verb())
  })

  it("returns verb from ppaction:// URL without query", {
    cNvPr <- make_cNvPr(make_hlink_click_xml(action = "ppaction://hlinkshowjump?jump=firstslide"))
    hlink <- cNvPr$hlinkClick
    expect_equal(hlink$action_verb(), "hlinkshowjump")
  })

  it("returns verb from ppaction:// URL without query string", {
    cNvPr <- make_cNvPr(make_hlink_click_xml(action = "ppaction://macro?name=MyMacro"))
    hlink <- cNvPr$hlinkClick
    expect_equal(hlink$action_verb(), "macro")
  })
})


describe("CT_Hyperlink$action_fields()", {
  it("returns empty list when no action attribute", {
    cNvPr <- make_cNvPr(make_hlink_click_xml())
    hlink <- cNvPr$hlinkClick
    expect_equal(hlink$action_fields(), list())
  })

  it("parses query string into named list", {
    cNvPr <- make_cNvPr(
      make_hlink_click_xml(action = "ppaction://hlinkshowjump?jump=firstslide")
    )
    hlink <- cNvPr$hlinkClick
    fields <- hlink$action_fields()
    expect_equal(fields[["jump"]], "firstslide")
  })

  it("parses multiple query string fields", {
    cNvPr <- make_cNvPr(
      make_hlink_click_xml(action = "ppaction://customshow?id=0&return=true")
    )
    hlink <- cNvPr$hlinkClick
    fields <- hlink$action_fields()
    expect_equal(fields[["id"]],     "0")
    expect_equal(fields[["return"]], "true")
  })
})


# ============================================================================
# ActionSetting$action — PP_ACTION_TYPE dispatch
# ============================================================================

# Simple parent stub with a NULL part (action read-path doesn't need part).
stub_parent <- list(part = NULL)

describe("ActionSetting$action", {
  it("returns NONE when no hlinkClick element", {
    cNvPr <- make_cNvPr()
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_identical(act$action, PP_ACTION_TYPE$NONE)
  })

  it("returns HYPERLINK when hlinkClick has no action attribute", {
    cNvPr <- make_cNvPr(make_hlink_click_xml())
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_identical(act$action, PP_ACTION_TYPE$HYPERLINK)
  })

  it("returns FIRST_SLIDE for hlinkshowjump?jump=firstslide", {
    cNvPr <- make_cNvPr(
      make_hlink_click_xml(action = "ppaction://hlinkshowjump?jump=firstslide")
    )
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_identical(act$action, PP_ACTION_TYPE$FIRST_SLIDE)
  })

  it("returns LAST_SLIDE for hlinkshowjump?jump=lastslide", {
    cNvPr <- make_cNvPr(
      make_hlink_click_xml(action = "ppaction://hlinkshowjump?jump=lastslide")
    )
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_identical(act$action, PP_ACTION_TYPE$LAST_SLIDE)
  })

  it("returns NEXT_SLIDE for hlinkshowjump?jump=nextslide", {
    cNvPr <- make_cNvPr(
      make_hlink_click_xml(action = "ppaction://hlinkshowjump?jump=nextslide")
    )
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_identical(act$action, PP_ACTION_TYPE$NEXT_SLIDE)
  })

  it("returns PREVIOUS_SLIDE for hlinkshowjump?jump=previousslide", {
    cNvPr <- make_cNvPr(
      make_hlink_click_xml(action = "ppaction://hlinkshowjump?jump=previousslide")
    )
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_identical(act$action, PP_ACTION_TYPE$PREVIOUS_SLIDE)
  })

  it("returns END_SHOW for hlinkshowjump?jump=endshow", {
    cNvPr <- make_cNvPr(
      make_hlink_click_xml(action = "ppaction://hlinkshowjump?jump=endshow")
    )
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_identical(act$action, PP_ACTION_TYPE$END_SHOW)
  })

  it("returns NAMED_SLIDE for hlinksldjump", {
    cNvPr <- make_cNvPr(
      make_hlink_click_xml(action = "ppaction://hlinksldjump")
    )
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_identical(act$action, PP_ACTION_TYPE$NAMED_SLIDE)
  })

  it("returns RUN_MACRO for macro verb", {
    cNvPr <- make_cNvPr(
      make_hlink_click_xml(action = "ppaction://macro?name=MyMacro")
    )
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_identical(act$action, PP_ACTION_TYPE$RUN_MACRO)
  })

  it("hover=TRUE reads hlinkHover instead of hlinkClick", {
    # hlinkHover with a hyperlink (no action) → HYPERLINK
    ns_a <- .nsmap[["a"]]
    ns_r <- .nsmap[["r"]]
    inner <- sprintf('<a:hlinkHover r:id="rId1"/>')
    cNvPr <- make_cNvPr(inner)
    act <- ActionSetting$new(cNvPr, stub_parent, hover = TRUE)
    expect_identical(act$action, PP_ACTION_TYPE$HYPERLINK)
  })
})


# ============================================================================
# ActionSetting$hyperlink — returns ShapeHyperlink
# ============================================================================

describe("ActionSetting$hyperlink", {
  it("returns a ShapeHyperlink", {
    cNvPr <- make_cNvPr(make_hlink_click_xml())
    act <- ActionSetting$new(cNvPr, stub_parent, hover = FALSE)
    expect_s3_class(act$hyperlink, "ShapeHyperlink")
  })

  it("returns a ShapeHyperlink for hover actions too", {
    cNvPr <- make_cNvPr()
    act <- ActionSetting$new(cNvPr, stub_parent, hover = TRUE)
    expect_s3_class(act$hyperlink, "ShapeHyperlink")
  })
})


# ============================================================================
# BaseShape$click_action and $hover_action
# ============================================================================

describe("BaseShape$click_action and $hover_action", {
  it("click_action returns ActionSetting with hover = FALSE", {
    prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    shape <- slide$shapes[[1]]
    act   <- shape$click_action
    expect_s3_class(act, "ActionSetting")
    expect_false(act$.__enclos_env__$private$.hover)
  })

  it("hover_action returns ActionSetting with hover = TRUE", {
    prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    shape <- slide$shapes[[1]]
    act   <- shape$hover_action
    expect_s3_class(act, "ActionSetting")
    expect_true(act$.__enclos_env__$private$.hover)
  })

  it("click_action$action is NONE for plain shape", {
    prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    shape <- slide$shapes[[1]]
    expect_identical(shape$click_action$action, PP_ACTION_TYPE$NONE)
  })
})
