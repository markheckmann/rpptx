# Base classes for proxy objects in rpptx.
#
# Ported from python-pptx/src/pptx/shared.py. Provides ElementProxy,
# ParentedElementProxy, and PartElementProxy which form the base of
# the domain object hierarchy.


# ============================================================================
# ElementProxy — wraps a single XML element
# ============================================================================

#' Base proxy wrapping an XML element
#'
#' @keywords internal
#' @export
ElementProxy <- R6::R6Class(
  "ElementProxy",

  public = list(
    initialize = function(element) {
      private$.element <- element
    }
  ),

  active = list(
    # The wrapped BaseOxmlElement
    element = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.element
    }
  ),

  private = list(
    .element = NULL
  )
)


# ============================================================================
# ParentedElementProxy — has a parent object
# ============================================================================

#' Proxy wrapping an XML element with a parent reference
#'
#' The parent is used to resolve the `part` property by delegation.
#'
#' @keywords internal
#' @export
ParentedElementProxy <- R6::R6Class(
  "ParentedElementProxy",
  inherit = ElementProxy,

  public = list(
    initialize = function(element, parent) {
      super$initialize(element)
      private$.parent <- parent
    }
  ),

  active = list(
    # The parent proxy object
    parent = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.parent
    },

    # The Part containing this element (resolved via parent chain)
    part = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.parent$part
    }
  ),

  private = list(
    .parent = NULL
  )
)


# ============================================================================
# PartElementProxy — wraps a part's root element
# ============================================================================

#' Proxy wrapping a part's root XML element
#'
#' Used for domain objects that correspond directly to an OPC part
#' (e.g. Presentation wraps the root element of PresentationPart).
#'
#' @keywords internal
#' @export
PartElementProxy <- R6::R6Class(
  "PartElementProxy",
  inherit = ElementProxy,

  public = list(
    initialize = function(element, part) {
      super$initialize(element)
      private$.part <- part
    }
  ),

  active = list(
    # The Part that contains this element
    part = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.part
    }
  ),

  private = list(
    .part = NULL
  )
)
