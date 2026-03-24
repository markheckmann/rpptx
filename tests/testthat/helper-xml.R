# Test helpers for XML operations.
#
# Provides utility functions for creating test XML elements, similar to
# python-pptx's tests/unitutil/cxml.py compact XML expression language.

#' Create an XML element from a string for testing
#'
#' @param xml_string A well-formed XML string.
#' @return An xml2 xml_document or xml_node.
#' @keywords internal
test_xml <- function(xml_string) {
  xml2::read_xml(xml_string)
}

#' Get the path to a test fixture file
#'
#' @param filename Name of the test fixture file.
#' @return Full path to the fixture file.
#' @keywords internal
test_file_path <- function(filename) {
  system.file("test_files", filename, package = "rpptx")
}

#' Get the path to a template file
#'
#' @param filename Name of the template file.
#' @return Full path to the template file.
#' @keywords internal
template_path <- function(filename = "default.pptx") {
  system.file("templates", filename, package = "rpptx")
}
