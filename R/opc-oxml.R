# OPC-local XML element classes for relationships and content types.
#
# Ported from python-pptx/src/pptx/opc/oxml.py. Handles OPC-specific XML
# elements: Relationships, Relationship, Types, Default, Override.


# ============================================================================
# Serialization helpers
# ============================================================================

#' Serialize a BaseOxmlElement to XML bytes
#'
#' @param element A BaseOxmlElement.
#' @return A raw vector of XML bytes.
#' @include oxml-init.R oxml-simpletypes.R
#' @noRd
serialize_part_xml <- function(element) {
  node <- if (inherits(element, "BaseOxmlElement")) element$get_node() else element
  xml_str <- as.character(node)
  # Strip any existing XML declaration before adding our own
  xml_str <- sub('^\\s*<\\?xml[^?]*\\?>\\s*', '', xml_str)
  xml_bytes <- charToRaw(paste0(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n', xml_str
  ))
  xml_bytes
}


# ============================================================================
# CT_Default — <Default> element
# ============================================================================

CT_Default <- define_oxml_element(

  classname = "CT_Default",
  tag = "ct:Default",
  attributes = list(
    extension = required_attribute("Extension", ST_Extension),
    contentType = required_attribute("ContentType", ST_ContentType)
  )
)


# ============================================================================
# CT_Override — <Override> element
# ============================================================================

CT_Override <- define_oxml_element(
  classname = "CT_Override",
  tag = "ct:Override",
  attributes = list(
    partName = required_attribute("PartName", XsdAnyUri),
    contentType = required_attribute("ContentType", ST_ContentType)
  )
)


# ============================================================================
# CT_Relationship — <Relationship> element
# ============================================================================

CT_Relationship <- define_oxml_element(
  classname = "CT_Relationship",
  tag = "pr:Relationship",
  attributes = list(
    rId = required_attribute("Id", XsdId),
    reltype = required_attribute("Type", XsdAnyUri),
    target_ref = required_attribute("Target", XsdAnyUri),
    targetMode = optional_attribute("TargetMode", ST_TargetMode, default = "Internal")
  ),
  methods = list(
    # Create a new <Relationship> element with attributes
    new_rel = function(rId, reltype, target_ref, target_mode = "Internal") {
      # Create the element manually with namespace
      ns_uri <- .nsmap[["pr"]]
      xml_str <- sprintf('<Relationship xmlns="%s"/>', ns_uri)
      doc <- xml2::read_xml(xml_str)
      node <- xml2::xml_root(doc)
      rel <- CT_Relationship$new(node)
      rel$rId <- rId
      rel$reltype <- reltype
      rel$target_ref <- target_ref
      rel$targetMode <- target_mode
      rel
    }
  )
)


# ============================================================================
# CT_Relationships — <Relationships> element (root of .rels files)
# ============================================================================

CT_Relationships <- define_oxml_element(
  classname = "CT_Relationships",
  tag = "pr:Relationships",
  children = list(
    relationship = zero_or_more("pr:Relationship")
  ),
  methods = list(

    # Add a child <Relationship> element
    add_rel = function(rId, reltype, target, is_external = FALSE) {
      target_mode <- if (is_external) "External" else "Internal"
      ns_uri <- .nsmap[["pr"]]
      xml_str <- sprintf('<Relationship xmlns="%s"/>', ns_uri)
      doc <- xml2::read_xml(xml_str)
      node <- xml2::xml_root(doc)
      rel <- CT_Relationship$new(node)
      rel$rId <- rId
      rel$reltype <- reltype
      rel$target_ref <- target
      rel$targetMode <- target_mode
      # Insert into this element
      self$`_insert_relationship`(rel)
      rel
    },

    # Get XML bytes suitable for saving in a .rels file
    xml_file_bytes = function() {
      serialize_part_xml(self)
    }
  )
)

#' Create a new empty <Relationships> element
#' @return A CT_Relationships instance.
#' @noRd
new_ct_relationships <- function() {
  ns_uri <- .nsmap[["pr"]]
  xml_str <- sprintf('<Relationships xmlns="%s"/>', ns_uri)
  doc <- xml2::read_xml(xml_str)
  node <- xml2::xml_root(doc)
  CT_Relationships$new(node)
}


# ============================================================================
# CT_Types — <Types> element (root of [Content_Types].xml)
# ============================================================================

CT_Types <- define_oxml_element(
  classname = "CT_Types",
  tag = "ct:Types",
  children = list(
    default = zero_or_more("ct:Default"),
    override = zero_or_more("ct:Override")
  ),
  methods = list(

    # Add a <Default> child element
    add_default = function(ext, content_type) {
      self$`_add_default`(extension = ext, contentType = content_type)
    },

    # Add an <Override> child element
    add_override = function(partname, content_type) {
      self$`_add_override`(partName = as.character(partname),
                           contentType = content_type)
    }
  )
)

#' Create a new empty <Types> element
#' @return A CT_Types instance.
#' @noRd
new_ct_types <- function() {
  ns_uri <- .nsmap[["ct"]]
  xml_str <- sprintf('<Types xmlns="%s"/>', ns_uri)
  doc <- xml2::read_xml(xml_str)
  node <- xml2::xml_root(doc)
  CT_Types$new(node)
}


# ============================================================================
# Register element classes
# ============================================================================

.onLoad_opc_oxml <- function() {
  register_element_cls("ct:Default", CT_Default)
  register_element_cls("ct:Override", CT_Override)
  register_element_cls("ct:Types", CT_Types)
  register_element_cls("pr:Relationship", CT_Relationship)
  register_element_cls("pr:Relationships", CT_Relationships)
}
