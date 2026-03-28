# Custom XML element classes for text-related XML elements.
#
# Ported from python-pptx/src/pptx/oxml/text.py.

#' @include utils.R oxml-xmlchemy.R oxml-simpletypes.R
#' @keywords internal


# ============================================================================
# CT_TextFont — <a:latin>, <a:ea>, <a:cs>, <a:sym>
# ============================================================================

#' @keywords internal
CT_TextFont <- define_oxml_element(
  classname = "CT_TextFont",
  tag = "a:latin",
  attributes = list(
    typeface = required_attribute("typeface", ST_TextTypeface)
  )
)


# ============================================================================
# CT_TextCharacterProperties — <a:rPr>, <a:defRPr>, <a:endParaRPr>
# ============================================================================

#' @keywords internal
CT_TextCharacterProperties <- define_oxml_element(
  classname = "CT_TextCharacterProperties",
  tag = "a:rPr",
  children = list(
    latin = zero_or_one(
      "a:latin",
      successors = c("a:ea", "a:cs", "a:sym", "a:hlinkClick",
                     "a:hlinkMouseOver", "a:rtl", "a:extLst")
    ),
    hlinkClick = zero_or_one(
      "a:hlinkClick",
      successors = c("a:hlinkMouseOver", "a:rtl", "a:extLst")
    )
  ),
  attributes = list(
    sz = optional_attribute("sz", ST_TextFontSize),
    b  = optional_attribute("b",  XsdBoolean),
    i  = optional_attribute("i",  XsdBoolean),
    u  = optional_attribute("u",  XsdString)
  )
)


# ============================================================================
# CT_Hyperlink — <a:hlinkClick>
# ============================================================================

#' CT_Hyperlink XML element
#' @keywords internal
#' @export
CT_Hyperlink <- R6::R6Class(
  "CT_Hyperlink",
  inherit = BaseOxmlElement,

  active = list(
    # The r:id attribute value (relationship ID) or NULL.
    # Uses BaseOxmlElement$set_attr/get_attr which convert Clark notation to
    # prefix:local with ns= — the correct xml2 API for namespaced attributes.
    rId = function(value) {
      clark <- paste0("{", .nsmap[["r"]], "}id")
      if (!missing(value)) {
        self$set_attr(clark, as.character(value))
        return(invisible(value))
      }
      self$get_attr(clark)
    }
  )
)


# ============================================================================
# CT_TextLineBreak — <a:br>
# ============================================================================

#' @keywords internal
CT_TextLineBreak <- define_oxml_element(
  classname = "CT_TextLineBreak",
  tag = "a:br",
  children = list(
    rPr = zero_or_one("a:rPr", successors = character(0))
  ),
  active = list(
    # Always returns a vertical-tab character (soft line break).
    text = function(value) {
      if (!missing(value)) stop("text is read-only for line breaks", call. = FALSE)
      "\v"
    }
  )
)


# ============================================================================
# CT_TextSpacingPercent — <a:spcPct>
# ============================================================================

#' @keywords internal
CT_TextSpacingPercent <- define_oxml_element(
  classname = "CT_TextSpacingPercent",
  tag = "a:spcPct",
  attributes = list(
    val = required_attribute("val", ST_TextSpacingPercentOrPercentString)
  )
)


# ============================================================================
# CT_TextSpacingPoint — <a:spcPts>
# ============================================================================

#' @keywords internal
CT_TextSpacingPoint <- define_oxml_element(
  classname = "CT_TextSpacingPoint",
  tag = "a:spcPts",
  attributes = list(
    val = required_attribute("val", ST_TextSpacingPoint)
  )
)


# ============================================================================
# CT_TextSpacing — <a:lnSpc>, <a:spcBef>, <a:spcAft>
# ============================================================================

#' @keywords internal
CT_TextSpacing <- define_oxml_element(
  classname = "CT_TextSpacing",
  tag = "a:lnSpc",
  children = list(
    spcPct = zero_or_one("a:spcPct"),
    spcPts = zero_or_one("a:spcPts")
  ),
  methods = list(
    # Set spacing to `value` lines (float); removes spcPts if present.
    set_spcPct = function(value) {
      self$`_remove_spcPts`()
      spcPct <- self$get_or_add_spcPct()
      spcPct$val <- value
    },
    # Set spacing to `value` EMU (Length); removes spcPct if present.
    set_spcPts = function(value) {
      self$`_remove_spcPct`()
      spcPts <- self$get_or_add_spcPts()
      spcPts$val <- value
    }
  )
)


# ============================================================================
# CT_TextParagraphProperties — <a:pPr>
# ============================================================================

#' @keywords internal
CT_TextParagraphProperties <- define_oxml_element(
  classname = "CT_TextParagraphProperties",
  tag = "a:pPr",
  children = list(
    lnSpc  = zero_or_one("a:lnSpc",   successors = c("a:spcBef", "a:spcAft", "a:defRPr", "a:extLst")),
    spcBef = zero_or_one("a:spcBef",  successors = c("a:spcAft", "a:defRPr", "a:extLst")),
    spcAft = zero_or_one("a:spcAft",  successors = c("a:defRPr", "a:extLst")),
    defRPr = zero_or_one("a:defRPr",  successors = c("a:extLst"))
  ),
  attributes = list(
    lvl  = optional_attribute("lvl",  ST_TextIndentLevelType, default = 0L),
    algn = optional_attribute("algn", XsdString)
  ),
  active = list(
    # Line spacing: float = lines, Length = fixed points, NULL = inherit.
    line_spacing = function(value) {
      if (!missing(value)) {
        self$`_remove_lnSpc`()
        if (!is.null(value)) {
          lnSpc <- self$`_add_lnSpc`()
          if (inherits(value, "Length")) lnSpc$set_spcPts(value) else lnSpc$set_spcPct(value)
        }
        return(invisible(value))
      }
      lnSpc <- self$lnSpc
      if (is.null(lnSpc)) return(NULL)
      if (!is.null(lnSpc$spcPts)) return(lnSpc$spcPts$val)
      if (!is.null(lnSpc$spcPct)) return(lnSpc$spcPct$val)
      NULL
    },
    # Space before paragraph in EMU (Length), or NULL.
    space_before = function(value) {
      if (!missing(value)) {
        self$`_remove_spcBef`()
        if (!is.null(value)) { spcBef <- self$`_add_spcBef`(); spcBef$set_spcPts(value) }
        return(invisible(value))
      }
      spcBef <- self$spcBef
      if (is.null(spcBef) || is.null(spcBef$spcPts)) return(NULL)
      spcBef$spcPts$val
    },
    # Space after paragraph in EMU (Length), or NULL.
    space_after = function(value) {
      if (!missing(value)) {
        self$`_remove_spcAft`()
        if (!is.null(value)) { spcAft <- self$`_add_spcAft`(); spcAft$set_spcPts(value) }
        return(invisible(value))
      }
      spcAft <- self$spcAft
      if (is.null(spcAft) || is.null(spcAft$spcPts)) return(NULL)
      spcAft$spcPts$val
    }
  )
)


# ============================================================================
# CT_RegularTextRun — <a:r>
# ============================================================================

#' @keywords internal
CT_RegularTextRun <- define_oxml_element(
  classname = "CT_RegularTextRun",
  tag = "a:r",
  children = list(
    rPr = zero_or_one("a:rPr",         successors = c("a:t")),
    t   = one_and_only_one("a:t")
  ),
  active = list(
    # Text content via the required <a:t> child (read/write).
    text = function(value) {
      if (!missing(value)) {
        xml2::xml_set_text(self$t$get_node(), as.character(value))
        return(invisible(value))
      }
      txt <- xml2::xml_text(self$t$get_node())
      if (is.na(txt) || is.null(txt)) "" else txt
    }
  )
)


# ============================================================================
# CT_TextParagraph — <a:p>
# ============================================================================

#' @keywords internal
CT_TextParagraph <- define_oxml_element(
  classname = "CT_TextParagraph",
  tag = "a:p",
  children = list(
    pPr        = zero_or_one("a:pPr",        successors = c("a:r", "a:br", "a:fld", "a:endParaRPr")),
    r          = zero_or_more("a:r",         successors = c("a:endParaRPr")),
    br         = zero_or_more("a:br",        successors = c("a:endParaRPr")),
    endParaRPr = zero_or_one("a:endParaRPr", successors = character(0))
  ),
  methods = list(
    # Append a new <a:r> child, optionally setting its text.
    add_r = function(text = NULL) {
      r <- self$`_add_r`()
      if (!is.null(text) && nchar(text) > 0L) r$text <- text
      r
    },
    # Append a new <a:br> line-break child.
    add_br = function() self$`_add_br`(),
    # Append runs and breaks for `text`, splitting on \n and \v.
    append_text = function(text) {
      parts <- strsplit(text, "\n|\v", perl = TRUE)[[1]]
      for (i in seq_along(parts)) {
        if (i > 1L) self$add_br()
        if (nchar(parts[i]) > 0L) self$add_r(parts[i])
      }
    }
  ),
  active = list(
    # List of a:r, a:br, a:fld children.
    content_children = function() {
      content_tags <- c(qn("a:r"), qn("a:br"), qn("a:fld"))
      result <- list()
      for (ch in xml2::xml_children(private$.node)) {
        if (.get_clark_name(ch) %in% content_tags) result <- c(result, list(wrap_element(ch)))
      }
      result
    },
    # All text in this paragraph (read/write). \v represents line breaks.
    text = function(value) {
      if (!missing(value)) {
        children <- self$content_children
        for (ch in children) xml2::xml_remove(ch$get_node())
        self$append_text(value)
        return(invisible(value))
      }
      paste0(vapply(self$content_children, function(e) {
        txt <- e$text
        if (is.null(txt) || (length(txt) == 1L && is.na(txt))) "" else as.character(txt)
      }, character(1)), collapse = "")
    }
  )
)

# Override auto-generated _new_r to produce <a:r><a:t/></a:r> (not bare <a:r/>).
CT_TextParagraph$set("public", "_new_r", function() {
  xmlns_a <- .nsmap[["a"]]
  xml_str <- sprintf('<a:r xmlns:a="%s"><a:t/></a:r>', xmlns_a)
  doc <- xml2::read_xml(xml_str)
  wrap_element(xml2::xml_root(doc))
}, overwrite = TRUE)


# ============================================================================
# CT_TextBodyProperties — <a:bodyPr>
# ============================================================================

#' @keywords internal
CT_TextBodyProperties <- define_oxml_element(
  classname = "CT_TextBodyProperties",
  tag = "a:bodyPr",
  attributes = list(
    lIns   = optional_attribute("lIns",   ST_Coordinate32, default = Emu(91440L)),
    tIns   = optional_attribute("tIns",   ST_Coordinate32, default = Emu(45720L)),
    rIns   = optional_attribute("rIns",   ST_Coordinate32, default = Emu(91440L)),
    bIns   = optional_attribute("bIns",   ST_Coordinate32, default = Emu(45720L)),
    anchor = optional_attribute("anchor", XsdString),
    wrap   = optional_attribute("wrap",   XsdString)
  )
)


# ============================================================================
# CT_TextBody — <p:txBody>, <a:txBody>
# ============================================================================

#' Custom element class for p:txBody and a:txBody elements.
#'
#' @keywords internal
#' @export
CT_TextBody <- R6::R6Class(
  "CT_TextBody",
  inherit = BaseOxmlElement,

  public = list(
    # Append a new empty <a:p> child and return it.
    add_p = function() {
      p <- OxmlElement("a:p")
      self$append_child(p)
      # Return wrapper pointing to the newly inserted node.
      self$p_lst[[length(self$p_lst)]]
    },

    # Remove all <a:p> children.
    clear_content = function() self$remove_all("a:p"),

    # Ensure at least one <a:p> child is present.
    unclear_content = function() {
      if (length(self$p_lst) == 0L) self$add_p()
    }
  ),

  active = list(
    # Required <a:bodyPr> child element.
    # No-op setter accepts the R6 write-back from txBody$bodyPr$prop <- val.
    bodyPr = function(value) {
      if (!missing(value)) return(invisible(NULL))
      child <- self$find(qn("a:bodyPr"))
      if (is.null(child)) stop("required <a:bodyPr> not found in txBody", call. = FALSE)
      child
    },

    # List of <a:p> children.
    p_lst = function() self$findall(qn("a:p")),

    # TRUE if contains only a single empty paragraph.
    is_empty = function() {
      ps <- self$p_lst
      if (length(ps) == 0L) return(TRUE)
      if (length(ps) > 1L)  return(FALSE)
      ps[[1]]$text == ""
    }
  )
)


# Factory: create a new <p:txBody> with <a:bodyPr/> and <a:p/> children.
.CT_TextBody_new_p_txBody <- function() {
  xmlns_p <- .nsmap[["p"]]
  xmlns_a <- .nsmap[["a"]]
  xml_str <- sprintf(
    '<p:txBody xmlns:p="%s" xmlns:a="%s"><a:bodyPr/><a:p/></p:txBody>',
    xmlns_p, xmlns_a
  )
  doc <- xml2::read_xml(xml_str)
  wrap_element(xml2::xml_root(doc))
}


# ============================================================================
# Register element classes
# ============================================================================

.onLoad_oxml_text <- function() {
  register_element_cls("a:latin",      CT_TextFont)
  register_element_cls("a:ea",         CT_TextFont)
  register_element_cls("a:cs",         CT_TextFont)
  register_element_cls("a:sym",        CT_TextFont)
  register_element_cls("a:rPr",        CT_TextCharacterProperties)
  register_element_cls("a:defRPr",     CT_TextCharacterProperties)
  register_element_cls("a:endParaRPr", CT_TextCharacterProperties)
  register_element_cls("a:br",         CT_TextLineBreak)
  register_element_cls("a:spcPct",     CT_TextSpacingPercent)
  register_element_cls("a:spcPts",     CT_TextSpacingPoint)
  register_element_cls("a:lnSpc",      CT_TextSpacing)
  register_element_cls("a:spcBef",     CT_TextSpacing)
  register_element_cls("a:spcAft",     CT_TextSpacing)
  register_element_cls("a:pPr",        CT_TextParagraphProperties)
  register_element_cls("a:p",          CT_TextParagraph)
  register_element_cls("a:r",          CT_RegularTextRun)
  register_element_cls("a:bodyPr",     CT_TextBodyProperties)
  register_element_cls("p:txBody",     CT_TextBody)
  register_element_cls("a:txBody",     CT_TextBody)
  register_element_cls("a:hlinkClick", CT_Hyperlink)
}
