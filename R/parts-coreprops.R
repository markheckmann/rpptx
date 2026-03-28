# CorePropertiesPart — OPC part for /docProps/core.xml
#
# Ported from python-pptx/src/pptx/parts/coreprops.py.

# ============================================================================
# CorePropertiesPart
# ============================================================================

#' Core properties part
#'
#' Corresponds to the `/docProps/core.xml` part in a .pptx package. Provides
#' read/write access to Dublin Core document metadata.
#'
#' @include opc-package.R oxml-coreprops.R
#' @keywords internal
#' @export
CorePropertiesPart <- R6::R6Class(
  "CorePropertiesPart",
  inherit = XmlPart,

  active = list(

    # Author (dc:creator)
    author = function(value) {
      if (!missing(value)) { self$element$author_text(value); return(invisible(value)) }
      self$element$author_text()
    },

    # Category (cp:category)
    category = function(value) {
      if (!missing(value)) { self$element$category_text(value); return(invisible(value)) }
      self$element$category_text()
    },

    # Comments / description (dc:description)
    comments = function(value) {
      if (!missing(value)) { self$element$comments_text(value); return(invisible(value)) }
      self$element$comments_text()
    },

    # Content status (cp:contentStatus)
    content_status = function(value) {
      if (!missing(value)) { self$element$contentStatus_text(value); return(invisible(value)) }
      self$element$contentStatus_text()
    },

    # Created datetime (dcterms:created)
    created = function(value) {
      if (!missing(value)) { self$element$created_datetime(value); return(invisible(value)) }
      self$element$created_datetime()
    },

    # Identifier (dc:identifier)
    identifier = function(value) {
      if (!missing(value)) { self$element$identifier_text(value); return(invisible(value)) }
      self$element$identifier_text()
    },

    # Keywords (cp:keywords)
    keywords = function(value) {
      if (!missing(value)) { self$element$keywords_text(value); return(invisible(value)) }
      self$element$keywords_text()
    },

    # Language (dc:language)
    language = function(value) {
      if (!missing(value)) { self$element$language_text(value); return(invisible(value)) }
      self$element$language_text()
    },

    # Last modified by (cp:lastModifiedBy)
    last_modified_by = function(value) {
      if (!missing(value)) { self$element$lastModifiedBy_text(value); return(invisible(value)) }
      self$element$lastModifiedBy_text()
    },

    # Last printed datetime (cp:lastPrinted)
    last_printed = function(value) {
      if (!missing(value)) { self$element$lastPrinted_datetime(value); return(invisible(value)) }
      self$element$lastPrinted_datetime()
    },

    # Modified datetime (dcterms:modified)
    modified = function(value) {
      if (!missing(value)) { self$element$modified_datetime(value); return(invisible(value)) }
      self$element$modified_datetime()
    },

    # Revision (cp:revision)
    revision = function(value) {
      if (!missing(value)) { self$element$revision_number(value); return(invisible(value)) }
      self$element$revision_number()
    },

    # Subject (dc:subject)
    subject = function(value) {
      if (!missing(value)) { self$element$subject_text(value); return(invisible(value)) }
      self$element$subject_text()
    },

    # Title (dc:title)
    title = function(value) {
      if (!missing(value)) { self$element$title_text(value); return(invisible(value)) }
      self$element$title_text()
    },

    # Version (cp:version)
    version = function(value) {
      if (!missing(value)) { self$element$version_text(value); return(invisible(value)) }
      self$element$version_text()
    }
  )
)


# ============================================================================
# CorePropertiesPart_default — create a new default instance
# ============================================================================

#' Create a default CorePropertiesPart
#' @keywords internal
CorePropertiesPart_default <- function(package) {
  core_props_elm <- new_ct_coreProperties()
  part <- CorePropertiesPart$new(
    PackURI("/docProps/core.xml"),
    CT$OPC_CORE_PROPERTIES,
    package,
    core_props_elm
  )
  part$title <- "PowerPoint Presentation"
  part$last_modified_by <- "rpptx"
  part$revision <- 1L
  part$modified <- Sys.time()
  part
}


# ============================================================================
# Register part type
# ============================================================================

.onLoad_parts_coreprops <- function() {
  register_part_type(CT$OPC_CORE_PROPERTIES, CorePropertiesPart)
}
