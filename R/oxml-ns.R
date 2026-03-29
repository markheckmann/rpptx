# Namespace related objects.
#
# Ported from python-pptx/src/pptx/oxml/ns.py. Provides namespace prefix
# mappings for all Office XML namespaces and utility functions for working
# with qualified (Clark notation) tag names.

# --- Namespace prefix to URI map ---

.nsmap <- list(
  a       = "http://schemas.openxmlformats.org/drawingml/2006/main",
  c       = "http://schemas.openxmlformats.org/drawingml/2006/chart",
  cp      = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  ct      = "http://schemas.openxmlformats.org/package/2006/content-types",
  dc      = "http://purl.org/dc/elements/1.1/",
  dcmitype = "http://purl.org/dc/dcmitype/",
  dcterms = "http://purl.org/dc/terms/",
  ep      = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  i       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  m       = "http://schemas.openxmlformats.org/officeDocument/2006/math",
  mo      = "http://schemas.microsoft.com/office/mac/office/2008/main",
  mv      = "urn:schemas-microsoft-com:mac:vml",
  o       = "urn:schemas-microsoft-com:office:office",
  p       = "http://schemas.openxmlformats.org/presentationml/2006/main",
  pd      = "http://schemas.openxmlformats.org/drawingml/2006/presentationDrawing",
  pic     = "http://schemas.openxmlformats.org/drawingml/2006/picture",
  pr      = "http://schemas.openxmlformats.org/package/2006/relationships",
  r       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  sl      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  v       = "urn:schemas-microsoft-com:vml",
  ve      = "http://schemas.openxmlformats.org/markup-compatibility/2006",
  w       = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  w10     = "urn:schemas-microsoft-com:office:word",
  wne     = "http://schemas.microsoft.com/office/word/2006/wordml",
  wp      = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  xsi     = "http://www.w3.org/2001/XMLSchema-instance"
)

# Reverse map: namespace URI -> prefix
.pfxmap <- stats::setNames(names(.nsmap), unlist(.nsmap, use.names = FALSE))


#' Get the Clark-notation qualified tag name for a namespace-prefixed tag
#'
#' Converts a namespace-prefixed tag like `"p:cSld"` to Clark notation like
#' `"{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"`.
#'
#' @param nsptag A namespace-prefixed tag string, e.g. `"p:cSld"`.
#' @return A string in Clark notation.
#' @export
#' @examples
#' qn("p:cSld")
#' qn("a:r")
qn <- function(nsptag) {
  parts <- strsplit(nsptag, ":", fixed = TRUE)[[1]]
  if (length(parts) != 2) {
    stop(sprintf("Invalid namespace-prefixed tag: '%s'", nsptag), call. = FALSE)
  }
  pfx <- parts[1]
  local_part <- parts[2]
  nsuri <- .nsmap[[pfx]]
  if (is.null(nsuri)) {
    stop(sprintf("Unknown namespace prefix: '%s'", pfx), call. = FALSE)
  }
  paste0("{", nsuri, "}", local_part)
}


#' Convert a Clark-notation tag to a namespace-prefixed tag
#'
#' @param clark_name A Clark-notation tag, e.g.
#'   `"{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"`.
#' @return A namespace-prefixed tag string, e.g. `"p:cSld"`.
#' @noRd
clark_to_nsptag <- function(clark_name) {
  # Parse "{uri}localname"
  m <- regmatches(clark_name, regexec("^\\{(.+?)\\}(.+)$", clark_name))[[1]]
  if (length(m) != 3) {
    stop(sprintf("Invalid Clark-notation tag: '%s'", clark_name), call. = FALSE)
  }
  nsuri <- m[2]
  local_part <- m[3]
  pfx <- .pfxmap[[nsuri]]
  if (is.null(pfx)) {
    stop(sprintf("Unknown namespace URI: '%s'", nsuri), call. = FALSE)
  }
  paste0(pfx, ":", local_part)
}


#' Get the namespace URI for a given prefix
#'
#' @param nspfx A namespace prefix string, e.g. `"p"`.
#' @return The namespace URI string.
#' @noRd
nsuri <- function(nspfx) {
  uri <- .nsmap[[nspfx]]
  if (is.null(uri)) {
    stop(sprintf("Unknown namespace prefix: '%s'", nspfx), call. = FALSE)
  }
  uri
}


#' Return a subset of the namespace map
#'
#' @param ... Namespace prefixes as character strings.
#' @return A named character vector mapping prefixes to URIs.
#' @noRd
namespaces <- function(...) {
  prefixes <- c(...)
  unlist(.nsmap[prefixes])
}


#' Return namespace declaration strings
#'
#' @param ... Namespace prefixes as character strings.
#' @return A single string with xmlns declarations, e.g.
#'   `'xmlns:p="http://..."'`.
#' @noRd
nsdecls <- function(...) {
  prefixes <- c(...)
  decls <- vapply(prefixes, function(pfx) {
    sprintf('xmlns:%s="%s"', pfx, .nsmap[[pfx]])
  }, character(1))
  paste(decls, collapse = " ")
}


#' Get the Clark name of an xml2 node
#'
#' xml2 nodes can report their tag in various ways. This function returns the
#' Clark-notation tag `{uri}localname` for use in the element registry.
#'
#' @param node An xml2 xml_node.
#' @return A string in Clark notation, or the plain tag name if no namespace.
#' @noRd
xml_clark_name <- function(node) {
  ns <- xml2::xml_ns(xml2::xml_root(node))
  name <- xml2::xml_name(node)

  # Check if name contains a prefix (e.g., "p:cSld")
  parts <- strsplit(name, ":", fixed = TRUE)[[1]]
  if (length(parts) == 2) {
    pfx <- parts[1]
    local <- parts[2]
    # Look up URI in the document's namespace map
    uri <- ns[[pfx]]
    if (!is.null(uri) && !is.na(uri)) {
      return(paste0("{", uri, "}", local))
    }
    # Try our own nsmap
    uri <- .nsmap[[pfx]]
    if (!is.null(uri)) {
      return(paste0("{", uri, "}", local))
    }
  }

  # If xml2 has already resolved the namespace, xml_name returns just the

  # local name and we need xml_ns() to find the default/element namespace
  # For elements with a default namespace, try to detect it
  name
}
