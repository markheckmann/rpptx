# XML parser initialization and element class registry.
#
# Ported from python-pptx/src/pptx/oxml/__init__.py. Provides the global
# element class registry, register_element_cls(), parse_xml(), and
# wrap_element() functions. In Python, lxml provides custom element class
# lookup automatically. In R, we implement this via an explicit registry.

# --- Global element class registry ---
# Maps Clark-notation tag names to R6 class generators
.element_registry <- new.env(parent = emptyenv())


#' Register an R6 class for a given XML element tag
#'
#' When the XML parser encounters an element with this tag, it will be wrapped
#' in an instance of `cls` instead of the default BaseOxmlElement.
#'
#' @param nsptagname Namespace-prefixed tag, e.g. `"p:presentation"`.
#' @param cls An R6ClassGenerator (e.g. the result of `define_oxml_element()`).
#' @include oxml-xmlchemy.R
#' @noRd
register_element_cls <- function(nsptagname, cls) {
  clark_name <- qn(nsptagname)
  .element_registry[[clark_name]] <- cls
}


#' Wrap an xml2 node in the appropriate R6 element class
#'
#' Looks up the node's tag in the element registry and returns an instance
#' of the registered class, or a BaseOxmlElement if no class is registered.
#'
#' @param node An xml2 xml_node.
#' @return A BaseOxmlElement (or subclass) instance.
#' @noRd
wrap_element <- function(node) {
  if (is.null(node) || inherits(node, "xml_missing")) return(NULL)
  # If it's already wrapped, return as-is

  if (inherits(node, "BaseOxmlElement")) return(node)

  clark_name <- .get_clark_name(node)
  cls <- .element_registry[[clark_name]]
  if (is.null(cls)) {
    return(BaseOxmlElement$new(node))
  }
  cls$new(node)
}


#' Parse XML bytes/string into a wrapped element
#'
#' @param xml A raw vector, character string, or connection containing XML.
#' @return A wrapped BaseOxmlElement (or registered subclass) for the root.
#' @noRd
rpptx_parse_xml <- function(xml) {
  if (is.raw(xml)) {
    doc <- xml2::read_xml(xml)
  } else {
    doc <- xml2::read_xml(xml)
  }
  root <- xml2::xml_root(doc)
  wrap_element(root)
}


#' Load and parse an XML template file
#'
#' @param template_name Name of the template (without .xml extension).
#' @return A wrapped element for the template root.
#' @noRd
parse_from_template <- function(template_name) {
  filename <- system.file("templates", paste0(template_name, ".xml"),
                          package = "rpptx")
  if (filename == "") {
    stop(sprintf("Template file not found: '%s.xml'", template_name), call. = FALSE)
  }
  xml_bytes <- readBin(filename, what = "raw", n = file.info(filename)$size)
  rpptx_parse_xml(xml_bytes)
}


# --- Internal helpers ---

#' Get the Clark-notation tag name for an xml2 node
#'
#' Uses xml2's xml_name() and namespace resolution to construct the Clark name.
#'
#' @param node An xml2 xml_node.
#' @return A Clark-notation tag string.
#' @noRd
.get_clark_name <- function(node) {
  # First try: xml_name with namespace context gives prefixed name
  doc_ns <- tryCatch(xml2::xml_ns(node), error = function(e) NULL)
  if (!is.null(doc_ns) && length(doc_ns) > 0) {
    ns_name <- xml2::xml_name(node, ns = doc_ns)
    parts <- strsplit(ns_name, ":", fixed = TRUE)[[1]]
    if (length(parts) == 2) {
      pfx <- parts[1]
      local <- parts[2]
      # Resolve prefix — try document namespace map first
      uri <- doc_ns[[pfx]]
      if (!is.null(uri) && !is.na(uri)) {
        return(paste0("{", uri, "}", local))
      }
      # Fall back to our known namespace map
      uri <- .nsmap[[pfx]]
      if (!is.null(uri)) {
        return(paste0("{", uri, "}", local))
      }
    }
  }

  # Fall back: raw name without namespace context
  raw_name <- xml2::xml_name(node)

  # Check if the raw name itself has a prefix
  parts <- strsplit(raw_name, ":", fixed = TRUE)[[1]]
  if (length(parts) == 2) {
    pfx <- parts[1]
    local <- parts[2]
    uri <- .nsmap[[pfx]]
    if (!is.null(uri)) {
      return(paste0("{", uri, "}", local))
    }
    if (!is.null(doc_ns)) {
      uri <- doc_ns[[pfx]]
      if (!is.null(uri) && !is.na(uri)) {
        return(paste0("{", uri, "}", local))
      }
    }
  }

  # No prefix at all — check for a single/default namespace
  if (!is.null(doc_ns) && length(doc_ns) > 0) {
    # For elements with a single namespace, use it
    if (length(doc_ns) == 1) {
      return(paste0("{", doc_ns[[1]], "}", raw_name))
    }
    # Check for default namespace (d1, d2, etc.)
    default_ns <- doc_ns[grepl("^d\\d+$", names(doc_ns))]
    if (length(default_ns) > 0) {
      return(paste0("{", default_ns[[1]], "}", raw_name))
    }
  }

  # No namespace at all
  raw_name
}
