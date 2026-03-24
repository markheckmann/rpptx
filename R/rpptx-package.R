#' @keywords internal
"_PACKAGE"

## usethis namespace: start
#' @importFrom R6 R6Class
#' @importFrom xml2 read_xml xml_find_all xml_find_first xml_text xml_attr
#'   xml_set_attr xml_name xml_ns xml_add_child xml_remove xml_children
#'   xml_root xml_new_root write_xml xml_set_text xml_add_sibling xml_parent
#'   xml_attrs xml_set_attrs as_xml_document xml_serialize
## usethis namespace: end
NULL

# Suppress R CMD check NOTEs for R6 standard variables
utils::globalVariables(c("self", "private", "super"))
