# Custom condition classes for rpptx.
#
# Ported from python-pptx/src/pptx/exc.py. Uses R's condition system
# with custom classes for structured error handling.

#' Signal an rpptx error
#'
#' @param message Error message.
#' @param class Character vector of additional subclasses.
#' @param call The call to include in the error.
#' @param ... Additional fields for the condition object.
#' @return No return value; raises a condition.
#' @noRd
rpptx_error <- function(message, class = NULL, call = NULL, ...) {
  stop(
    structure(
      class = c(class, "rpptx_error", "error", "condition"),
      list(message = message, call = call, ...)
    )
  )
}

#' Signal a "package not found" error
#'
#' Raised when a .pptx package cannot be found at the specified path.
#'
#' @param path The file path that was not found.
#' @noRd
package_not_found_error <- function(path) {
  rpptx_error(
    message = sprintf("Package not found: '%s'", path),
    class = "rpptx_package_not_found_error",
    path = path
  )
}

#' Signal an "invalid XML" error
#'
#' Raised when a value in the XML is not valid according to the schema.
#'
#' @param message Description of the invalid XML.
#' @noRd
invalid_xml_error <- function(message) {
  rpptx_error(
    message = message,
    class = "rpptx_invalid_xml_error"
  )
}
