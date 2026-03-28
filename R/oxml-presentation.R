# Custom element classes for presentation-related XML elements.
#
# Ported from python-pptx/src/pptx/oxml/presentation.py.

# ============================================================================
# CT_SlideSize — <p:sldSz>
# ============================================================================

#' @include oxml-xmlchemy.R oxml-simpletypes.R oxml-init.R
#' @keywords internal
CT_SlideSize <- define_oxml_element(
  classname = "CT_SlideSize",
  tag = "p:sldSz",
  attributes = list(
    cx = required_attribute("cx", ST_SlideSizeCoordinate),
    cy = required_attribute("cy", ST_SlideSizeCoordinate)
  )
)


# ============================================================================
# CT_SlideId — <p:sldId>
# ============================================================================

#' @keywords internal
CT_SlideId <- define_oxml_element(
  classname = "CT_SlideId",
  tag = "p:sldId",
  attributes = list(
    id = required_attribute("id", ST_SlideId),
    rId = required_attribute("r:id", XsdString)
  )
)


# ============================================================================
# CT_SlideIdList — <p:sldIdLst>
# ============================================================================

#' @keywords internal
CT_SlideIdList <- define_oxml_element(
  classname = "CT_SlideIdList",
  tag = "p:sldIdLst",
  children = list(
    sldId = zero_or_more("p:sldId")
  ),
  methods = list(
    # Add a new sldId child with the given rId
    add_sldId = function(rId) {
      next_id <- private$.next_id()
      self$`_add_sldId`(id = next_id, rId = rId)
    }
  )
)

# Attach private _next_id method (needs access to self)
CT_SlideIdList$set("private", ".next_id", function() {
  MIN_SLIDE_ID <- 256L
  MAX_SLIDE_ID <- 2147483647L

  # Get all used IDs from child p:sldId elements
  used_ids <- integer(0)
  for (sld_id in self$sldId_lst) {
    used_ids <- c(used_ids, sld_id$id)
  }

  simple_next <- max(c(MIN_SLIDE_ID - 1L, used_ids)) + 1L
  if (simple_next <= MAX_SLIDE_ID) {
    return(simple_next)
  }

  # Fall back to search for next unused from bottom
  valid_used <- sort(used_ids[used_ids >= MIN_SLIDE_ID & used_ids <= MAX_SLIDE_ID])
  if (length(valid_used) == 0) return(MIN_SLIDE_ID)

  for (i in seq_along(valid_used)) {
    candidate <- MIN_SLIDE_ID + i - 1L
    if (candidate != valid_used[i]) return(candidate)
  }
  MIN_SLIDE_ID + length(valid_used)
})


# ============================================================================
# CT_SlideMasterIdListEntry — <p:sldMasterId>
# ============================================================================

#' @keywords internal
CT_SlideMasterIdListEntry <- define_oxml_element(
  classname = "CT_SlideMasterIdListEntry",
  tag = "p:sldMasterId",
  attributes = list(
    rId = required_attribute("r:id", XsdString)
  )
)


# ============================================================================
# CT_SlideMasterIdList — <p:sldMasterIdLst>
# ============================================================================

#' @keywords internal
CT_SlideMasterIdList <- define_oxml_element(
  classname = "CT_SlideMasterIdList",
  tag = "p:sldMasterIdLst",
  children = list(
    sldMasterId = zero_or_more("p:sldMasterId")
  )
)


# ============================================================================
# CT_Presentation — <p:presentation>
# ============================================================================

#' @keywords internal
CT_Presentation <- define_oxml_element(
  classname = "CT_Presentation",
  tag = "p:presentation",
  children = list(
    sldMasterIdLst = zero_or_one(
      "p:sldMasterIdLst",
      successors = c("p:notesMasterIdLst", "p:handoutMasterIdLst",
                     "p:sldIdLst", "p:sldSz", "p:notesSz")
    ),
    sldIdLst = zero_or_one(
      "p:sldIdLst",
      successors = c("p:sldSz", "p:notesSz")
    ),
    sldSz = zero_or_one(
      "p:sldSz",
      successors = c("p:notesSz")
    )
  )
)


# ============================================================================
# Register element classes
# ============================================================================

.onLoad_oxml_presentation <- function() {
  register_element_cls("p:presentation", CT_Presentation)
  register_element_cls("p:sldSz", CT_SlideSize)
  register_element_cls("p:sldId", CT_SlideId)
  register_element_cls("p:sldIdLst", CT_SlideIdList)
  register_element_cls("p:sldMasterId", CT_SlideMasterIdListEntry)
  register_element_cls("p:sldMasterIdLst", CT_SlideMasterIdList)
}
