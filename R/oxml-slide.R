# Custom element classes for slide-related XML elements.
#
# Ported from python-pptx/src/pptx/oxml/slide.py.

# ============================================================================
# CT_CommonSlideData — <p:cSld>
# ============================================================================

#' @include oxml-xmlchemy.R oxml-simpletypes.R oxml-init.R
#' @keywords internal
CT_CommonSlideData <- define_oxml_element(
  classname = "CT_CommonSlideData",
  tag = "p:cSld",
  children = list(
    bg = zero_or_one(
      "p:bg",
      successors = c("p:spTree", "p:custDataLst", "p:controls", "p:extLst")
    ),
    spTree = one_and_only_one("p:spTree")
  ),
  attributes = list(
    name = optional_attribute("name", XsdString, default = "")
  )
)


# ============================================================================
# _BaseSlideElement helpers — spTree shortcut via cSld
# ============================================================================
# Each slide-type element class (CT_Slide, CT_SlideLayout, CT_SlideMaster)
# has a cSld child (OneAndOnlyOne) and exposes spTree/bg via delegation.

.add_slide_element_methods <- function(cls) {
  cls$set("active", "spTree", function(value) {
    if (!missing(value)) stop("Read-only", call. = FALSE)
    self$cSld$spTree
  })
  cls$set("active", "bg", function(value) {
    if (!missing(value)) stop("Read-only", call. = FALSE)
    self$cSld$bg
  })
  cls
}


# ============================================================================
# CT_Slide — <p:sld>
# ============================================================================

#' @keywords internal
CT_Slide <- define_oxml_element(
  classname = "CT_Slide",
  tag = "p:sld",
  children = list(
    cSld = one_and_only_one("p:cSld"),
    clrMapOvr = zero_or_one(
      "p:clrMapOvr",
      successors = c("p:transition", "p:timing", "p:extLst")
    ),
    timing = zero_or_one(
      "p:timing",
      successors = c("p:extLst")
    )
  )
)
CT_Slide <- .add_slide_element_methods(CT_Slide)


# Factory: create new blank slide element
new_ct_slide <- function() {
  xml_str <- sprintf(
    paste0(
      '<p:sld %s>\n',
      '  <p:cSld>\n',
      '    <p:spTree>\n',
      '      <p:nvGrpSpPr>\n',
      '        <p:cNvPr id="1" name=""/>\n',
      '        <p:cNvGrpSpPr/>\n',
      '        <p:nvPr/>\n',
      '      </p:nvGrpSpPr>\n',
      '      <p:grpSpPr/>\n',
      '    </p:spTree>\n',
      '  </p:cSld>\n',
      '  <p:clrMapOvr>\n',
      '    <a:masterClrMapping/>\n',
      '  </p:clrMapOvr>\n',
      '</p:sld>'
    ),
    nsdecls("a", "p", "r")
  )
  rpptx_parse_xml(charToRaw(xml_str))
}


# ============================================================================
# CT_SlideLayout — <p:sldLayout>
# ============================================================================

#' @keywords internal
CT_SlideLayout <- define_oxml_element(
  classname = "CT_SlideLayout",
  tag = "p:sldLayout",
  children = list(
    cSld = one_and_only_one("p:cSld")
  )
)
CT_SlideLayout <- .add_slide_element_methods(CT_SlideLayout)


# ============================================================================
# CT_SlideLayoutIdListEntry — <p:sldLayoutId>
# ============================================================================

#' @keywords internal
CT_SlideLayoutIdListEntry <- define_oxml_element(
  classname = "CT_SlideLayoutIdListEntry",
  tag = "p:sldLayoutId",
  attributes = list(
    rId = required_attribute("r:id", XsdString)
  )
)


# ============================================================================
# CT_SlideLayoutIdList — <p:sldLayoutIdLst>
# ============================================================================

#' @keywords internal
CT_SlideLayoutIdList <- define_oxml_element(
  classname = "CT_SlideLayoutIdList",
  tag = "p:sldLayoutIdLst",
  children = list(
    sldLayoutId = zero_or_more("p:sldLayoutId")
  )
)


# ============================================================================
# CT_SlideMaster — <p:sldMaster>
# ============================================================================

#' @keywords internal
CT_SlideMaster <- define_oxml_element(
  classname = "CT_SlideMaster",
  tag = "p:sldMaster",
  children = list(
    cSld = one_and_only_one("p:cSld"),
    sldLayoutIdLst = zero_or_one(
      "p:sldLayoutIdLst",
      successors = c("p:transition", "p:timing", "p:hf", "p:txStyles", "p:extLst")
    )
  )
)
CT_SlideMaster <- .add_slide_element_methods(CT_SlideMaster)


# ============================================================================
# Register element classes
# ============================================================================

.onLoad_oxml_slide <- function() {
  register_element_cls("p:cSld", CT_CommonSlideData)
  register_element_cls("p:sld", CT_Slide)
  register_element_cls("p:sldLayout", CT_SlideLayout)
  register_element_cls("p:sldLayoutId", CT_SlideLayoutIdListEntry)
  register_element_cls("p:sldLayoutIdLst", CT_SlideLayoutIdList)
  register_element_cls("p:sldMaster", CT_SlideMaster)
}
