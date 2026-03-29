# Custom element classes for slide-related XML elements.
#
# Ported from python-pptx/src/pptx/oxml/slide.py.

# ============================================================================
# CT_CommonSlideData — <p:cSld>
# ============================================================================

#' @include oxml-xmlchemy.R oxml-simpletypes.R oxml-init.R
#' @noRd
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

# Return <p:bg/p:bgPr> grandchild, creating a noFill background if absent.
CT_CommonSlideData$set("public", "get_or_add_bgPr", function() {
  bg <- self$bg
  if (is.null(bg) || is.null(wrap_element(bg$get_node())$bgPr)) {
    bg <- private$.change_to_noFill_bg()
  }
  bg <- self$bg
  if (inherits(bg, "BaseOxmlElement")) bg <- wrap_element(bg$get_node())
  bgPr_nd <- xml2::xml_find_first(bg$get_node(), "p:bgPr",
                                   ns = c(p = .nsmap[["p"]]))
  wrap_element(bgPr_nd)
})

# Remove existing bg and add a noFill bgPr; return the new CT_Background.
CT_CommonSlideData$set("private", ".change_to_noFill_bg", function() {
  existing <- xml2::xml_find_first(self$get_node(), "p:bg",
                                    ns = c(p = .nsmap[["p"]]))
  if (!inherits(existing, "xml_missing")) xml2::xml_remove(existing)
  bg_nd <- xml2::xml_add_child(
    self$get_node(), "p:bg", .where = 0L,  # insert before spTree
    xmlns = .nsmap[["p"]]
  )
  bgPr_nd <- xml2::xml_add_child(bg_nd, "p:bgPr", xmlns = .nsmap[["p"]])
  xml2::xml_add_child(bgPr_nd, "a:noFill", xmlns = .nsmap[["a"]])
  xml2::xml_add_child(bgPr_nd, "a:effectLst", xmlns = .nsmap[["a"]])
  wrap_element(bg_nd)
})


# ============================================================================
# CT_Background — <p:bg>
# ============================================================================

#' @noRd
CT_Background <- R6::R6Class(
  "CT_Background",
  inherit = BaseOxmlElement,

  active = list(
    # <p:bgPr> child or NULL
    bgPr = function() {
      nd <- xml2::xml_find_first(self$get_node(), "p:bgPr",
                                 ns = c(p = .nsmap[["p"]]))
      if (inherits(nd, "xml_missing")) return(NULL)
      wrap_element(nd)
    }
  )
)


# ============================================================================
# CT_BackgroundProperties — <p:bgPr>
# ============================================================================

#' @noRd
CT_BackgroundProperties <- R6::R6Class(
  "CT_BackgroundProperties",
  inherit = BaseOxmlElement
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

#' @noRd
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

#' @noRd
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

#' @noRd
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

#' @noRd
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

#' @noRd
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
# CT_NotesSlide — <p:notes>
# ============================================================================

#' @noRd
CT_NotesSlide <- define_oxml_element(
  classname = "CT_NotesSlide",
  tag = "p:notes",
  children = list(
    cSld = one_and_only_one("p:cSld"),
    clrMapOvr = zero_or_one(
      "p:clrMapOvr",
      successors = c("p:extLst")
    )
  )
)
CT_NotesSlide <- .add_slide_element_methods(CT_NotesSlide)


# ============================================================================
# CT_NotesMaster — <p:notesMaster>
# ============================================================================

#' @noRd
CT_NotesMaster <- define_oxml_element(
  classname = "CT_NotesMaster",
  tag = "p:notesMaster",
  children = list(
    cSld = one_and_only_one("p:cSld")
  )
)
CT_NotesMaster <- .add_slide_element_methods(CT_NotesMaster)


# ============================================================================
# Register element classes
# ============================================================================

.onLoad_oxml_slide <- function() {
  register_element_cls("p:cSld",          CT_CommonSlideData)
  register_element_cls("p:bg",            CT_Background)
  register_element_cls("p:bgPr",          CT_BackgroundProperties)
  register_element_cls("p:sld",           CT_Slide)
  register_element_cls("p:sldLayout",     CT_SlideLayout)
  register_element_cls("p:sldLayoutId",   CT_SlideLayoutIdListEntry)
  register_element_cls("p:sldLayoutIdLst",CT_SlideLayoutIdList)
  register_element_cls("p:sldMaster",     CT_SlideMaster)
  register_element_cls("p:notes",         CT_NotesSlide)
  register_element_cls("p:notesMaster",   CT_NotesMaster)
}
