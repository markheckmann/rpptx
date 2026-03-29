# Core properties XML element.
#
# Ported from python-pptx/src/pptx/oxml/coreprops.py.

# ============================================================================
# CT_CoreProperties — cp:coreProperties
# ============================================================================

#' CT_CoreProperties XML element
#' @noRd
CT_CoreProperties <- R6::R6Class(
  "CT_CoreProperties",
  inherit = BaseOxmlElement,

  public = list(

    # --- text properties ---

    # Get/set text of dc:creator (author)
    author_text = function(value) {
      if (!missing(value)) { private$.set_text("dc", "creator", value); return(invisible(value)) }
      private$.text_of("dc", "creator")
    },

    # Get/set text of cp:category
    category_text = function(value) {
      if (!missing(value)) { private$.set_text("cp", "category", value); return(invisible(value)) }
      private$.text_of("cp", "category")
    },

    # Get/set text of dc:description (comments)
    comments_text = function(value) {
      if (!missing(value)) { private$.set_text("dc", "description", value); return(invisible(value)) }
      private$.text_of("dc", "description")
    },

    # Get/set text of cp:contentStatus
    contentStatus_text = function(value) {
      if (!missing(value)) { private$.set_text("cp", "contentStatus", value); return(invisible(value)) }
      private$.text_of("cp", "contentStatus")
    },

    # Get/set text of dc:identifier
    identifier_text = function(value) {
      if (!missing(value)) { private$.set_text("dc", "identifier", value); return(invisible(value)) }
      private$.text_of("dc", "identifier")
    },

    # Get/set text of cp:keywords
    keywords_text = function(value) {
      if (!missing(value)) { private$.set_text("cp", "keywords", value); return(invisible(value)) }
      private$.text_of("cp", "keywords")
    },

    # Get/set text of dc:language
    language_text = function(value) {
      if (!missing(value)) { private$.set_text("dc", "language", value); return(invisible(value)) }
      private$.text_of("dc", "language")
    },

    # Get/set text of cp:lastModifiedBy
    lastModifiedBy_text = function(value) {
      if (!missing(value)) { private$.set_text("cp", "lastModifiedBy", value); return(invisible(value)) }
      private$.text_of("cp", "lastModifiedBy")
    },

    # Get/set text of dc:subject
    subject_text = function(value) {
      if (!missing(value)) { private$.set_text("dc", "subject", value); return(invisible(value)) }
      private$.text_of("dc", "subject")
    },

    # Get/set text of dc:title
    title_text = function(value) {
      if (!missing(value)) { private$.set_text("dc", "title", value); return(invisible(value)) }
      private$.text_of("dc", "title")
    },

    # Get/set text of cp:version
    version_text = function(value) {
      if (!missing(value)) { private$.set_text("cp", "version", value); return(invisible(value)) }
      private$.text_of("cp", "version")
    },

    # --- datetime properties ---

    # Get/set dcterms:created as POSIXct or NULL
    created_datetime = function(value) {
      if (!missing(value)) { private$.set_datetime("dcterms", "created", value); return(invisible(value)) }
      private$.datetime_of("dcterms", "created")
    },

    # Get/set cp:lastPrinted as POSIXct or NULL
    lastPrinted_datetime = function(value) {
      if (!missing(value)) { private$.set_datetime("cp", "lastPrinted", value); return(invisible(value)) }
      private$.datetime_of("cp", "lastPrinted")
    },

    # Get/set dcterms:modified as POSIXct or NULL
    modified_datetime = function(value) {
      if (!missing(value)) { private$.set_datetime("dcterms", "modified", value); return(invisible(value)) }
      private$.datetime_of("dcterms", "modified")
    },

    # --- revision (integer) ---

    # Get integer value of cp:revision, or 0 if missing/invalid
    revision_number = function(value) {
      if (!missing(value)) {
        if (!is.numeric(value) || value < 1 || value != as.integer(value)) {
          stop("revision must be a positive integer", call. = FALSE)
        }
        elm <- private$.get_or_add_child("cp", "revision")
        xml2::xml_set_text(elm, as.character(as.integer(value)))
        return(invisible(value))
      }
      elm <- private$.find_child("cp", "revision")
      if (is.null(elm)) return(0L)
      txt <- xml2::xml_text(elm)
      n <- suppressWarnings(as.integer(txt))
      if (is.na(n) || n < 1L) return(0L)
      n
    }
  ),

  private = list(

    # Return text content of first matching child element, or ""
    .text_of = function(ns_pfx, local_name) {
      elm <- private$.find_child(ns_pfx, local_name)
      if (is.null(elm)) return("")
      txt <- xml2::xml_text(elm)
      if (is.na(txt) || is.null(txt)) "" else txt
    },

    # Set text content of named child element (creating if needed), truncate to 255 chars
    .set_text = function(ns_pfx, local_name, value) {
      value <- as.character(value)
      if (nchar(value) > 255L) {
        stop(sprintf("exceeded 255 char limit for '%s:%s'", ns_pfx, local_name), call. = FALSE)
      }
      elm <- private$.get_or_add_child(ns_pfx, local_name)
      xml2::xml_set_text(elm, value)
    },

    # Return POSIXct datetime from W3CDTF string in named child, or NULL
    .datetime_of = function(ns_pfx, local_name) {
      elm <- private$.find_child(ns_pfx, local_name)
      if (is.null(elm)) return(NULL)
      txt <- xml2::xml_text(elm)
      if (is.na(txt) || nchar(txt) == 0L) return(NULL)
      .parse_W3CDTF(txt)
    },

    # Set W3CDTF datetime string in named child element (creating if needed)
    .set_datetime = function(ns_pfx, local_name, value) {
      if (!inherits(value, c("POSIXct", "POSIXlt", "Date"))) {
        stop("datetime value must be a POSIXct, POSIXlt, or Date", call. = FALSE)
      }
      dt_str <- format(as.POSIXct(value, tz = "UTC"), "%Y-%m-%dT%H:%M:%SZ")
      elm <- private$.get_or_add_child(ns_pfx, local_name)
      xml2::xml_set_text(elm, dt_str)
      # created and modified require xsi:type="dcterms:W3CDTF"
      if (local_name %in% c("created", "modified")) {
        xml2::xml_set_attr(elm, paste0("{", .nsmap[["xsi"]], "}type"), "dcterms:W3CDTF")
      }
    },

    # Return the first matching child element by ns prefix + local name, or NULL
    .find_child = function(ns_pfx, local_name) {
      ns_uri <- .nsmap[[ns_pfx]]
      clark  <- sprintf("{%s}%s", ns_uri, local_name)
      node   <- private$.node
      children <- xml2::xml_children(node)
      for (ch in children) {
        tag <- .get_clark_name(ch)
        if (identical(tag, clark)) return(ch)
      }
      NULL
    },

    # Return or create a child element with the given ns prefix + local name
    .get_or_add_child = function(ns_pfx, local_name) {
      existing <- private$.find_child(ns_pfx, local_name)
      if (!is.null(existing)) return(existing)
      ns_uri <- .nsmap[[ns_pfx]]
      new_xml <- sprintf('<%s:%s xmlns:%s="%s"/>', ns_pfx, local_name, ns_pfx, ns_uri)
      new_node <- xml2::xml_root(xml2::read_xml(new_xml))
      xml2::xml_add_child(private$.node, new_node)
      # Re-find to get the node in context
      private$.find_child(ns_pfx, local_name)
    }
  )
)


# ============================================================================
# Helper: parse W3CDTF datetime string to POSIXct
# ============================================================================

.parse_W3CDTF <- function(w3cdtf_str) {
  # Strip trailing offset like +05:00 or -07:30 (not handled by strptime)
  # "Z" suffix = UTC
  parseable <- substr(w3cdtf_str, 1L, 19L)
  offset    <- substr(w3cdtf_str, 20L, nchar(w3cdtf_str))

  formats <- c("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%Y-%m", "%Y")
  ts <- NULL
  for (fmt in formats) {
    ts <- suppressWarnings(as.POSIXct(parseable, format = fmt, tz = "UTC"))
    if (!is.na(ts)) break
  }
  if (is.null(ts) || is.na(ts)) return(NULL)

  # Apply numeric timezone offset if present
  if (nchar(offset) == 6L && substr(offset, 1L, 1L) %in% c("+", "-")) {
    sign    <- if (substr(offset, 1L, 1L) == "+") -1L else 1L  # offset is FROM UTC, we go TO UTC
    h_off   <- as.integer(substr(offset, 2L, 3L))
    m_off   <- as.integer(substr(offset, 5L, 6L))
    ts <- ts + sign * (h_off * 3600L + m_off * 60L)
  }
  ts
}


# ============================================================================
# Factory: build a new empty cp:coreProperties element
# ============================================================================

#' Create a new empty cp:coreProperties element
#' @noRd
new_ct_coreProperties <- function() {
  xml_str <- sprintf(
    '<cp:coreProperties %s/>',
    nsdecls("cp", "dc", "dcterms", "xsi")
  )
  rpptx_parse_xml(xml_str)
}


# ============================================================================
# Register element class
# ============================================================================

.onLoad_oxml_coreprops <- function() {
  register_element_cls("cp:coreProperties", CT_CoreProperties)
}
