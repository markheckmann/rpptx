# Fundamental Open Packaging Convention (OPC) objects.
#
# Ported from python-pptx/src/pptx/opc/package.py. Provides OpcPackage, Part,
# XmlPart, PartFactory, Relationships, and related classes for reading and
# writing presentations to and from a .pptx file.


# ============================================================================
# OpcPackage — Main API class for OPC packages
# ============================================================================

#' Open Packaging Convention package
#'
#' A new instance is constructed by calling `OpcPackage$open()` with a path
#' to a package file (.pptx).
#'
#' @include opc-oxml.R opc-serialized.R opc-constants.R opc-packuri.R opc-spec.R
#' @noRd
#' @export
OpcPackage <- R6::R6Class(
  "OpcPackage",

  public = list(
    initialize = function(pkg_file) {
      private$.pkg_file <- pkg_file
    },

    # Drop relationship identified by `rId`
    drop_rel = function(rId) {
      self$rels$pop(rId)
    },

    # Generate exactly one reference to each part in the package
    # @return A list of Part objects.
    iter_parts = function() {
      visited <- list()
      parts <- list()
      for (rel in self$iter_rels()) {
        if (rel$is_external) next
        part <- rel$target_part
        # Use partname as identity key (R doesn't have Python's id() / set())
        pn <- as.character(part$partname)
        if (pn %in% names(visited)) next
        visited[[pn]] <- TRUE
        parts[[length(parts) + 1]] <- part
      }
      parts
    },

    # Generate exactly one reference to each relationship in package
    # Depth-first traversal of the rels graph.
    # @return A list of Relationship objects.
    iter_rels = function() {
      visited <- list()
      result <- list()

      walk_rels <- function(rels) {
        for (rel in rels$values()) {
          result[[length(result) + 1]] <<- rel
          if (rel$is_external) next
          part <- rel$target_part
          pn <- as.character(part$partname)
          if (pn %in% names(visited)) next
          visited[[pn]] <<- TRUE
          walk_rels(part$rels)
        }
      }

      walk_rels(self$rels)
      result
    },

    # Return the next available partname matching template `tmpl`
    #
    # @param tmpl A sprintf-style template with one `%d`, e.g.
    #   `"/ppt/slides/slide%d.xml"`.
    # @return A PackURI.
    next_partname = function(tmpl) {
      # Find the prefix (everything before the number)
      prefix <- sub("%d.*$", "", tmpl)
      partnames <- character(0)
      for (p in self$iter_parts()) {
        pn <- as.character(p$partname)
        if (startsWith(pn, prefix)) {
          partnames <- c(partnames, pn)
        }
      }
      for (n in seq(length(partnames) + 1, 1, by = -1)) {
        candidate <- sprintf(tmpl, n)
        if (!(candidate %in% partnames)) {
          return(PackURI(candidate))
        }
      }
      stop("ProgrammingError: ran out of candidate_partnames", call. = FALSE)
    },

    # Return Part having relationship to this package of `reltype`
    # @param reltype Relationship type string.
    # @return A Part.
    part_related_by = function(reltype) {
      self$rels$part_with_reltype(reltype)
    },

    # Return rId of relationship of `reltype` to `target`
    # @param target A Part object or string (for external).
    # @param reltype Relationship type string.
    # @param is_external Logical.
    # @return A string rId.
    relate_to = function(target, reltype, is_external = FALSE) {
      if (is.character(target)) {
        return(self$rels$get_or_add_ext_rel(reltype, target))
      }
      self$rels$get_or_add(reltype, target)
    },

    # Return related Part identified by `rId`
    related_part = function(rId) {
      self$rels$get(rId)$target_part
    },

    # Return URL in target ref of relationship identified by `rId`
    target_ref = function(rId) {
      self$rels$get(rId)$target_ref
    },

    # Save this package to `pkg_file`
    save = function(pkg_file) {
      writer <- PackageWriter$new()
      writer$write(pkg_file, self$rels, self$iter_parts())
    }
  ),

  active = list(
    # Main document part (PresentationPart)
    main_document_part = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      self$part_related_by(RT$OFFICE_DOCUMENT)
    },

    # Relationships collection for this package
    rels = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      if (is.null(private$.rels)) {
        private$.rels <- Relationships$new(pack_uri_base(PACKAGE_URI))
      }
      private$.rels
    }
  ),

  private = list(
    .pkg_file = NULL,
    .rels = NULL,

    # Load package contents from file
    .load = function() {
      result <- PackageLoader$new(private$.pkg_file, self)$load()
      pkg_xml_rels <- result$pkg_xml_rels
      parts <- result$parts
      self$rels$load_from_xml(pack_uri_base(PACKAGE_URI), pkg_xml_rels, parts)
      invisible(self)
    }
  )
)

#' Open an OPC package file
#' @param pkg_file Path to a .pptx file.
#' @return An OpcPackage instance.
#' @noRd
#' @export
OpcPackage_open <- function(pkg_file) {
  pkg <- OpcPackage$new(pkg_file)
  pkg$.__enclos_env__$private$.load()
  pkg
}


# ============================================================================
# PackageLoader — loads a package from disk
# ============================================================================

#' @noRd
PackageLoader <- R6::R6Class(
  "PackageLoader",

  public = list(
    initialize = function(pkg_file, package) {
      private$.pkg_file <- pkg_file
      private$.package <- package
    },

    # Load package and return list(pkg_xml_rels, parts)
    load = function() {
      parts <- private$.parts()
      xml_rels <- private$.xml_rels()

      for (partname in names(parts)) {
        part <- parts[[partname]]
        rels_for_part <- xml_rels[[partname]]
        if (!is.null(rels_for_part)) {
          part$load_rels_from_xml(rels_for_part, parts)
        }
      }

      list(
        pkg_xml_rels = xml_rels[[as.character(PACKAGE_URI)]],
        parts = parts
      )
    }
  ),

  private = list(
    .pkg_file = NULL,
    .package = NULL,
    .content_types_cache = NULL,
    .package_reader_cache = NULL,
    .parts_cache = NULL,
    .xml_rels_cache = NULL,

    .content_types = function() {
      if (is.null(private$.content_types_cache)) {
        pr <- private$.package_reader()
        ct_blob <- pr$get_blob(CONTENT_TYPES_URI)
        private$.content_types_cache <- ContentTypeMap$from_xml(ct_blob)
      }
      private$.content_types_cache
    },

    .package_reader = function() {
      if (is.null(private$.package_reader_cache)) {
        private$.package_reader_cache <- PackageReader$new(private$.pkg_file)
      }
      private$.package_reader_cache
    },

    .parts = function() {
      if (!is.null(private$.parts_cache)) return(private$.parts_cache)

      content_types <- private$.content_types()
      package <- private$.package
      package_reader <- private$.package_reader()
      xml_rels <- private$.xml_rels()

      parts <- list()
      for (partname_str in names(xml_rels)) {
        if (partname_str == "/") next
        partname <- PackURI(partname_str)
        # Skip invalid partnames not present in the package
        if (!package_reader$contains(partname)) next

        ct <- content_types$get(partname)
        blob <- package_reader$get_blob(partname)
        part <- PartFactory_create(partname, ct, package, blob)
        parts[[partname_str]] <- part
      }

      private$.parts_cache <- parts
      parts
    },

    .xml_rels = function() {
      if (!is.null(private$.xml_rels_cache)) return(private$.xml_rels_cache)

      xml_rels <- list()
      visited <- character(0)

      load_rels <- function(source_partname, rels) {
        key <- as.character(source_partname)
        xml_rels[[key]] <<- rels
        visited <<- c(visited, key)
        base_uri <- pack_uri_base(source_partname)

        for (rel in rels$relationship_lst) {
          if (rel$targetMode == RTM$EXTERNAL) next
          target_partname <- pack_uri_from_rel_ref(base_uri, rel$target_ref)
          target_key <- as.character(target_partname)
          if (target_key %in% visited) next
          load_rels(target_partname, private$.xml_rels_for(target_partname))
        }
      }

      load_rels(PACKAGE_URI, private$.xml_rels_for(PACKAGE_URI))
      private$.xml_rels_cache <- xml_rels
      xml_rels
    },

    .xml_rels_for = function(partname) {
      pr <- private$.package_reader()
      rels_xml <- pr$rels_xml_for(partname)
      if (is.null(rels_xml)) {
        return(new_ct_relationships())
      }
      rpptx_parse_xml(rels_xml)
    }
  )
)


# ============================================================================
# Part — Base class for package parts
# ============================================================================

#' Base class for package parts
#'
#' Provides common properties and methods for all part types.
#'
#' @noRd
#' @export
Part <- R6::R6Class(
  "Part",

  public = list(
    initialize = function(partname, content_type, package, blob = NULL) {
      private$.partname <- partname
      private$.content_type <- content_type
      private$.package <- package
      private$.blob <- blob
    },

    # Load relationships from parsed XML
    load_rels_from_xml = function(xml_rels, parts) {
      self$rels$load_from_xml(
        pack_uri_base(private$.partname), xml_rels, parts
      )
    },

    # Check if this part has any relationships
    has_rels = function() {
      length(self$rels) > 0
    },

    # Get XML bytes for this part's relationships
    rels_xml_bytes = function() {
      self$rels$xml_bytes()
    },

    # Return rId of relationship of `reltype` to `target`
    relate_to = function(target, reltype, is_external = FALSE) {
      if (is.character(target)) {
        return(self$rels$get_or_add_ext_rel(reltype, target))
      }
      self$rels$get_or_add(reltype, target)
    },

    # Return Part having relationship to this part of `reltype`
    part_related_by = function(reltype) {
      self$rels$part_with_reltype(reltype)
    },

    # Return related Part identified by `rId`
    related_part = function(rId) {
      self$rels$get(rId)$target_part
    },

    # Return target ref string for relationship identified by `rId`
    target_ref = function(rId) {
      self$rels$get(rId)$target_ref
    },

    # Drop relationship identified by `rId`
    drop_rel = function(rId) {
      self$rels$pop(rId)
    }
  ),

  active = list(
    # Raw blob bytes of this part
    blob = function(value) {
      if (!missing(value)) {
        private$.blob <- value
        return(invisible(value))
      }
      private$.blob %||% raw(0)
    },

    # Content type (MIME type) of this part
    content_type = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.content_type
    },

    # The package this part belongs to
    package = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.package
    },

    # PackURI partname for this part
    partname = function(value) {
      if (!missing(value)) {
        if (!inherits(value, "PackURI")) {
          stop(sprintf(
            "partname must be a PackURI, got '%s'",
            class(value)[1]
          ), call. = FALSE)
        }
        private$.partname <- value
        return(invisible(value))
      }
      private$.partname
    },

    # Relationships collection for this part
    rels = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      if (is.null(private$.rels)) {
        private$.rels <- Relationships$new(pack_uri_base(private$.partname))
      }
      private$.rels
    }
  ),

  private = list(
    .partname = NULL,
    .content_type = NULL,
    .package = NULL,
    .blob = NULL,
    .rels = NULL
  )
)


# ============================================================================
# XmlPart — Base class for XML-containing parts
# ============================================================================

#' Base class for parts containing an XML payload
#'
#' Provides additional methods for parsing/reserializing XML and managing
#' relationships. Most package parts are XmlParts.
#'
#' @noRd
#' @export
XmlPart <- R6::R6Class(
  "XmlPart",
  inherit = Part,

  public = list(
    initialize = function(partname, content_type, package, element) {
      super$initialize(partname, content_type, package)
      private$.element <- element
    },

    # Drop relationship if reference count < 2
    drop_rel = function(rId) {
      if (private$.rel_ref_count(rId) < 2) {
        self$rels$pop(rId)
      }
    }
  ),

  active = list(
    # XML bytes serialization of this part
    blob = function(value) {
      if (!missing(value)) {
        stop("XmlPart blob is read-only (derived from element)", call. = FALSE)
      }
      serialize_part_xml(private$.element)
    },

    # The XML element for this part
    element = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.element
    },

    # This part (part of parent protocol)
    part = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      self
    }
  ),

  private = list(
    .element = NULL,

    .rel_ref_count = function(rId) {
      # Count references to rId in this part's XML
      node <- if (inherits(private$.element, "BaseOxmlElement")) {
        private$.element$get_node()
      } else {
        private$.element
      }
      matches <- xml2::xml_find_all(node, paste0("//@r:id"),
        ns = c(r = .nsmap[["r"]]))
      sum(xml2::xml_text(matches) == rId)
    }
  )
)

#' Load an XmlPart from a blob
#' @param cls The R6 class generator to use.
#' @param partname A PackURI.
#' @param content_type Content type string.
#' @param package The package.
#' @param blob Raw bytes of XML.
#' @return An XmlPart instance.
#' @noRd
XmlPart_load <- function(cls, partname, content_type, package, blob) {
  element <- rpptx_parse_xml(blob)
  cls$new(partname, content_type, package, element)
}


# ============================================================================
# PartFactory — dispatches part construction by content type
# ============================================================================

# Global registry mapping content_type -> Part R6 class
.part_type_registry <- new.env(parent = emptyenv())

#' Register a Part subclass for a content type
#' @param content_type Content type string.
#' @param part_cls An R6 class generator.
#' @noRd
#' @export
register_part_type <- function(content_type, part_cls) {
  .part_type_registry[[content_type]] <- part_cls
}

#' Construct the appropriate Part subclass for a given content type
#' @param partname A PackURI.
#' @param content_type Content type string.
#' @param package The package.
#' @param blob Raw bytes.
#' @return A Part instance.
#' @noRd
#' @export
PartFactory_create <- function(partname, content_type, package, blob) {
  part_cls <- .part_type_registry[[content_type]]
  if (is.null(part_cls)) {
    # Default: use Part for binary or XmlPart for XML content
    if (grepl("+xml$", content_type, fixed = FALSE) ||
        content_type == CT$XML) {
      return(XmlPart_load(XmlPart, partname, content_type, package, blob))
    }
    return(Part$new(partname, content_type, package, blob))
  }
  # Registered classes are expected to have a $load() classmethod-style
  if (inherits(part_cls, "R6ClassGenerator") &&
      "load" %in% names(part_cls$public_methods)) {
    return(part_cls$public_methods$load(partname, content_type, package, blob))
  }
  # Fallback: try XmlPart_load with the class
  XmlPart_load(part_cls, partname, content_type, package, blob)
}


# ============================================================================
# ContentTypeMap — resolves partname -> content type
# ============================================================================

#' Resolves partname to content type from \[Content_Types\].xml
#' @noRd
ContentTypeMap <- R6::R6Class(
  "ContentTypeMap",

  public = list(
    initialize = function(overrides, defaults) {
      private$.overrides <- overrides
      private$.defaults <- defaults
    },

    # Get content type for a partname
    # @param partname A PackURI.
    # @return Content type string.
    get = function(partname) {
      key <- tolower(as.character(partname))
      if (key %in% names(private$.overrides)) {
        return(private$.overrides[[key]])
      }
      ext <- tolower(pack_uri_ext(partname))
      if (ext %in% names(private$.defaults)) {
        return(private$.defaults[[ext]])
      }
      stop(sprintf(
        "no content-type for partname '%s' in [Content_Types].xml",
        as.character(partname)
      ), call. = FALSE)
    }
  ),

  private = list(
    .overrides = NULL,
    .defaults = NULL
  )
)

#' Create a ContentTypeMap from \[Content_Types\].xml bytes
#' @param content_types_xml Raw bytes of \[Content_Types\].xml.
#' @return A ContentTypeMap instance.
#' @noRd
ContentTypeMap$from_xml <- function(content_types_xml) {
  types_elm <- rpptx_parse_xml(content_types_xml)

  overrides <- list()
  for (o in types_elm$override_lst) {
    key <- tolower(o$partName)
    overrides[[key]] <- o$contentType
  }

  defaults <- list()
  for (d in types_elm$default_lst) {
    key <- tolower(d$extension)
    defaults[[key]] <- d$contentType
  }

  ContentTypeMap$new(overrides, defaults)
}


# ============================================================================
# Relationships — collection of Relationship objects
# ============================================================================

#' Collection of relationships from a part or package to other parts
#'
#' Keyed by rId. Supports dict-like access.
#'
#' @noRd
#' @export
Relationships <- R6::R6Class(
  "Relationships",

  public = list(
    initialize = function(base_uri) {
      private$.base_uri <- base_uri
      private$.rels <- list()
    },

    # Get a relationship by rId
    # @param rId Relationship ID string.
    # @return A Relationship object.
    get = function(rId) {
      rel <- private$.rels[[rId]]
      if (is.null(rel)) {
        stop(sprintf("no relationship with key '%s'", rId), call. = FALSE)
      }
      rel
    },

    # Check if an rId exists
    contains = function(rId) {
      rId %in% names(private$.rels)
    },

    # Get all relationship values
    values = function() {
      unname(private$.rels)
    },

    # Get all rId keys
    keys = function() {
      names(private$.rels)
    },

    # Get or add a relationship to a target part
    # @param reltype Relationship type string.
    # @param target_part A Part object.
    # @return The rId string.
    get_or_add = function(reltype, target_part) {
      existing <- private$.get_matching(reltype, target_part)
      if (!is.null(existing)) return(existing)
      private$.add_relationship(reltype, target_part)
    },

    # Get or add an external relationship
    # @param reltype Relationship type string.
    # @param target_ref A URL string.
    # @return The rId string.
    get_or_add_ext_rel = function(reltype, target_ref) {
      existing <- private$.get_matching(reltype, target_ref, is_external = TRUE)
      if (!is.null(existing)) return(existing)
      private$.add_relationship(reltype, target_ref, is_external = TRUE)
    },

    # Load relationships from parsed XML
    # @param base_uri Base URI for resolving relative references.
    # @param xml_rels A CT_Relationships element.
    # @param parts Named list of parts by partname.
    load_from_xml = function(base_uri, xml_rels, parts) {
      private$.rels <- list()
      for (rel_elm in xml_rels$relationship_lst) {
        # Skip broken relationships (e.g. pointing to NULL)
        if (rel_elm$targetMode == RTM$INTERNAL) {
          partname <- pack_uri_from_rel_ref(base_uri, rel_elm$target_ref)
          if (!(as.character(partname) %in% names(parts))) next
        }
        rel <- Relationship_from_xml(base_uri, rel_elm, parts)
        private$.rels[[rel$rId]] <- rel
      }
    },

    # Return target part of relationship with matching `reltype`
    # @param reltype Relationship type string.
    # @return A Part.
    part_with_reltype = function(reltype) {
      matching <- list()
      for (rel in private$.rels) {
        if (rel$reltype == reltype) {
          matching[[length(matching) + 1]] <- rel
        }
      }
      if (length(matching) == 0) {
        stop(sprintf(
          "no relationship of type '%s' in collection", reltype
        ), call. = FALSE)
      }
      if (length(matching) > 1) {
        stop(sprintf(
          "multiple relationships of type '%s' in collection", reltype
        ), call. = FALSE)
      }
      matching[[1]]$target_part
    },

    # Remove and return relationship identified by `rId`
    pop = function(rId) {
      rel <- private$.rels[[rId]]
      private$.rels[[rId]] <- NULL
      rel
    },

    # Get XML bytes for serialization as a .rels file
    xml_bytes = function() {
      rels_elm <- new_ct_relationships()

      # Sort relationships by numerical rId order
      rIds <- names(private$.rels)
      nums <- vapply(rIds, function(rId) {
        m <- regmatches(rId, regexec("^rId([0-9]+)$", rId))[[1]]
        if (length(m) < 2) 0L else as.integer(m[2])
      }, integer(1))
      sorted_rIds <- rIds[order(nums)]

      for (rId in sorted_rIds) {
        rel <- private$.rels[[rId]]
        rels_elm$add_rel(rel$rId, rel$reltype, rel$target_ref, rel$is_external)
      }

      rels_elm$xml_file_bytes()
    }
  ),

  private = list(
    .base_uri = NULL,
    .rels = NULL,

    .add_relationship = function(reltype, target, is_external = FALSE) {
      rId <- private$.next_rId()
      target_mode <- if (is_external) RTM$EXTERNAL else RTM$INTERNAL
      private$.rels[[rId]] <- Relationship$new(
        private$.base_uri, rId, reltype, target_mode, target
      )
      rId
    },

    .get_matching = function(reltype, target, is_external = FALSE) {
      rels_of_type <- list()
      for (rel in private$.rels) {
        if (rel$reltype == reltype) {
          rels_of_type[[length(rels_of_type) + 1]] <- rel
        }
      }
      for (rel in rels_of_type) {
        if (rel$is_external != is_external) next
        rel_target <- if (rel$is_external) rel$target_ref else rel$target_part
        if (identical(rel_target, target)) return(rel$rId)
      }
      NULL
    },

    .next_rId = function() {
      n_rels <- length(private$.rels)
      for (n in seq(n_rels + 1, 1, by = -1)) {
        candidate <- sprintf("rId%d", n)
        if (!(candidate %in% names(private$.rels))) {
          return(candidate)
        }
      }
      stop("ProgrammingError: impossible rId collision", call. = FALSE)
    }
  )
)

# S3 method for length() on Relationships
#' @export
length.Relationships <- function(x) {
  length(x$.__enclos_env__$private$.rels)
}


# ============================================================================
# Relationship — single relationship value object
# ============================================================================

#' Value object describing a link from a part or package to another part
#' @noRd
#' @export
Relationship <- R6::R6Class(
  "Relationship",
  cloneable = FALSE,

  public = list(
    initialize = function(base_uri, rId, reltype, target_mode, target) {
      private$.base_uri <- base_uri
      private$.rId <- rId
      private$.reltype <- reltype
      private$.target_mode <- target_mode
      private$.target <- target
    }
  ),

  active = list(
    # TRUE if this is an external relationship
    is_external = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.target_mode == RTM$EXTERNAL
    },

    # Relationship type
    reltype = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.reltype
    },

    # Relationship ID (e.g. "rId1")
    rId = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      private$.rId
    },

    # Target Part (raises error if external)
    target_part = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      if (self$is_external) {
        stop("target_part is undefined for external relationships",
             call. = FALSE)
      }
      private$.target
    },

    # Target partname (raises error if external)
    target_partname = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      if (self$is_external) {
        stop("target_partname is undefined for external relationships",
             call. = FALSE)
      }
      private$.target$partname
    },

    # Target reference string (relative partname or URL)
    target_ref = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      if (self$is_external) {
        return(private$.target)
      }
      pack_uri_relative_ref(private$.target$partname, private$.base_uri)
    }
  ),

  private = list(
    .base_uri = NULL,
    .rId = NULL,
    .reltype = NULL,
    .target_mode = NULL,
    .target = NULL
  )
)

#' Create a Relationship from a CT_Relationship XML element
#' @param base_uri Base URI string.
#' @param rel A CT_Relationship element.
#' @param parts Named list mapping partname -> Part.
#' @return A Relationship object.
#' @noRd
Relationship_from_xml <- function(base_uri, rel, parts) {
  if (rel$targetMode == RTM$EXTERNAL) {
    target <- rel$target_ref
  } else {
    partname <- pack_uri_from_rel_ref(base_uri, rel$target_ref)
    target <- parts[[as.character(partname)]]
  }
  Relationship$new(base_uri, rel$rId, rel$reltype, rel$targetMode, target)
}
