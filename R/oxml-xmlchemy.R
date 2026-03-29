# Base classes enabling declarative definition of custom XML element classes.
#
# Ported from python-pptx/src/pptx/oxml/xmlchemy.py. Provides BaseOxmlElement
# R6 class and the define_oxml_element() factory that generates R6 classes
# with active bindings for child elements and attributes.


# ============================================================================
# BaseOxmlElement — R6 class wrapping an xml2 node
# ============================================================================

#' Base class for all custom XML element classes
#'
#' Wraps an xml2 xml_node and provides standardized methods for child element
#' access, insertion, removal, and attribute manipulation.
#'
#' @include oxml-ns.R
#' @noRd
#' @export
BaseOxmlElement <- R6::R6Class(
  "BaseOxmlElement",

  public = list(

    # Create a new element wrapper.
    # @param node An xml2 xml_node.
    initialize = function(node) {
      private$.node <- node
    },

    # Get the underlying xml2 node.
    get_node = function() {
      private$.node
    },

    # Find the first child element matching a Clark-notation tag.
    # @param clark_name Tag in Clark notation, e.g. `"{uri}localname"`.
    # @return A wrapped element or NULL.
    find = function(clark_name) {
      # Use xpath with namespace map to find child
      nsptag <- clark_to_nsptag(clark_name)
      parts <- strsplit(nsptag, ":", fixed = TRUE)[[1]]
      pfx <- parts[1]
      ns <- stats::setNames(.nsmap[[pfx]], pfx)
      result <- xml2::xml_find_first(private$.node, nsptag, ns = ns)
      if (inherits(result, "xml_missing")) return(NULL)
      wrap_element(result)
    },

    # Find all child elements matching a Clark-notation tag.
    # @param clark_name Tag in Clark notation.
    # @return A list of wrapped elements (possibly empty).
    findall = function(clark_name) {
      nsptag <- clark_to_nsptag(clark_name)
      parts <- strsplit(nsptag, ":", fixed = TRUE)[[1]]
      pfx <- parts[1]
      ns <- stats::setNames(.nsmap[[pfx]], pfx)
      results <- xml2::xml_find_all(private$.node, nsptag, ns = ns)
      lapply(results, wrap_element)
    },

    # Get an attribute value.
    # @param attr_name Clark-notation name ({uri}local) or plain local name.
    # @return The attribute value as a string, or NULL.
    get_attr = function(attr_name) {
      # Clark notation: {uri}localname — use namespace-aware lookup
      m <- regmatches(attr_name, regexec("^\\{([^}]+)\\}(.+)$", attr_name))[[1]]
      if (length(m) == 3) {
        uri <- m[2]; local <- m[3]
        # Build a namespace map for xml2 lookup
        pfx <- .pfxmap[[uri]]
        if (!is.null(pfx)) {
          ns_vec <- c(pfx, uri); names(ns_vec) <- NULL
          ns_vec <- setNames(uri, pfx)
          val <- xml2::xml_attr(private$.node,
                                paste0(pfx, ":", local),
                                ns = ns_vec)
          if (!is.na(val)) return(val)
        }
        # Fallback: scan xml_attrs for matching Clark name
        all_attrs <- xml2::xml_attrs(private$.node)
        ns_map <- tryCatch(xml2::xml_ns(private$.node), error = function(e) NULL)
        for (aname in names(all_attrs)) {
          parts <- strsplit(aname, ":", fixed = TRUE)[[1]]
          if (length(parts) == 2) {
            a_pfx <- parts[1]; a_local <- parts[2]
            a_uri <- if (!is.null(ns_map)) ns_map[[a_pfx]] else .nsmap[[a_pfx]]
            if (!is.null(a_uri) && !is.na(a_uri) && a_uri == uri && a_local == local) {
              return(all_attrs[[aname]])
            }
          }
        }
        return(NULL)
      }
      # Plain name (no namespace)
      val <- xml2::xml_attr(private$.node, attr_name)
      if (is.na(val)) NULL else val
    },

    # Set an attribute value.
    # @param attr_name Clark-notation name ({uri}local) or plain local name.
    # @param value The value to set (character string).
    set_attr = function(attr_name, value) {
      # Clark notation: map to prefixed name for xml2
      m <- regmatches(attr_name, regexec("^\\{([^}]+)\\}(.+)$", attr_name))[[1]]
      if (length(m) == 3) {
        uri <- m[2]; local <- m[3]
        pfx <- .pfxmap[[uri]]
        if (!is.null(pfx)) {
          ns_vec <- setNames(uri, pfx)
          xml2::xml_set_attr(private$.node, paste0(pfx, ":", local),
                             value, ns = ns_vec)
          return(invisible(self))
        }
      }
      xml2::xml_set_attr(private$.node, attr_name, value)
      invisible(self)
    },

    # Remove an attribute. xml2::xml_set_attr(node, name, NULL) removes it.
    remove_attr = function(attr_name) {
      # Extract local name from Clark notation if needed
      m <- regmatches(attr_name, regexec("^\\{([^}]+)\\}(.+)$", attr_name))[[1]]
      local_name <- if (length(m) == 3) m[3] else attr_name
      xml2::xml_set_attr(private$.node, local_name, NULL)
      invisible(self)
    },

    # Get all attributes as a named character vector.
    get_attrs = function() {
      xml2::xml_attrs(private$.node)
    },

    # Append a child element node.
    # @param child A BaseOxmlElement or xml2 xml_node.
    # Returns the child (wrapper updated to point to the inserted node).
    append_child = function(child) {
      child_node <- if (inherits(child, "BaseOxmlElement")) child$get_node() else child
      xml2::xml_add_child(private$.node, child_node)
      # xml_add_child copies cross-document nodes; re-find inserted node
      if (inherits(child, "BaseOxmlElement")) {
        children <- xml2::xml_children(private$.node)
        child$.__enclos_env__$private$.node <- children[[length(children)]]
      }
      invisible(self)
    },

    # Remove a child element node.
    # @param child A BaseOxmlElement or xml2 xml_node.
    remove_child = function(child) {
      child_node <- if (inherits(child, "BaseOxmlElement")) child$get_node() else child
      xml2::xml_remove(child_node)
      invisible(self)
    },

    # First child with tag in `tagnames`, or NULL if not found.
    # @param ... Namespace-prefixed tagnames, e.g. `"a:p"`, `"a:r"`.
    first_child_found_in = function(...) {
      tagnames <- c(...)
      for (tagname in tagnames) {
        child <- self$find(qn(tagname))
        if (!is.null(child)) return(child)
      }
      NULL
    },

    # Insert `elm` before the first child matching any of `tagnames`.
    # @param elm A BaseOxmlElement to insert.
    # @param ... Namespace-prefixed tagnames of successor elements.
    # @return The inserted element (wrapper updated to point to the inserted node).
    insert_element_before = function(elm, ...) {
      successor <- self$first_child_found_in(...)
      elm_node <- if (inherits(elm, "BaseOxmlElement")) elm$get_node() else elm
      if (!is.null(successor)) {
        succ_node <- if (inherits(successor, "BaseOxmlElement")) {
          successor$get_node()
        } else {
          successor
        }
        # xml_add_sibling returns the inserted node (in the parent doc)
        inserted_node <- xml2::xml_add_sibling(succ_node, elm_node, .where = "before")
        if (inherits(elm, "BaseOxmlElement")) {
          elm$.__enclos_env__$private$.node <- inserted_node
        }
      } else {
        xml2::xml_add_child(private$.node, elm_node)
        # xml_add_child returns parent; re-find the last child
        if (inherits(elm, "BaseOxmlElement")) {
          children <- xml2::xml_children(private$.node)
          elm$.__enclos_env__$private$.node <- children[[length(children)]]
        }
      }
      elm
    },

    # Remove all child elements with the given tagnames.
    # @param ... Namespace-prefixed tagnames, e.g. `"a:p"`.
    remove_all = function(...) {
      tagnames <- c(...)
      for (tagname in tagnames) {
        children <- self$findall(qn(tagname))
        for (child in children) {
          xml2::xml_remove(child$get_node())
        }
      }
      invisible(self)
    },

    # Run an xpath query with the standard namespace map.
    # @param xpath_str XPath expression string.
    # @return An xml2 nodeset.
    xpath = function(xpath_str) {
      ns <- unlist(.nsmap)
      xml2::xml_find_all(private$.node, xpath_str, ns = ns)
    },

    # Get pretty-printed XML string.
    # @return Character string of XML.
    to_xml = function() {
      as.character(private$.node)
    }
  ),

  active = list(

    # The Clark-notation tag name of this element.
    tag = function() {
      xml_clark_name(private$.node)
    },

    # Pretty-printed XML string (read-only).
    xml = function(value) {
      if (!missing(value)) stop("xml is read-only", call. = FALSE)
      self$to_xml()
    },

    # The text content of this element.
    text = function(value) {
      if (missing(value)) {
        xml2::xml_text(private$.node)
      } else {
        xml2::xml_set_text(private$.node, value)
      }
    }
  ),

  private = list(
    .node = NULL
  )
)


# ============================================================================
# OxmlElement() — create a new "loose" element
# ============================================================================

#' Create a new XML element with the given namespace-prefixed tag
#'
#' @param nsptag_str Namespace-prefixed tag, e.g. `"a:tbl"`.
#' @param nsmap Optional named character vector of additional namespace mappings.
#' @return A wrapped BaseOxmlElement (or appropriate subclass).
#' @noRd
#' @export
OxmlElement <- function(nsptag_str, nsmap = NULL) {
  parts <- strsplit(nsptag_str, ":", fixed = TRUE)[[1]]
  pfx <- parts[1]
  local_part <- parts[2]
  uri <- .nsmap[[pfx]]
  if (is.null(uri)) {
    stop(sprintf("Unknown namespace prefix: '%s'", pfx), call. = FALSE)
  }

  # Build XML string with namespace declaration
  xml_str <- sprintf('<%s xmlns:%s="%s"/>', nsptag_str, pfx, uri)
  doc <- xml2::read_xml(xml_str)
  node <- xml2::xml_root(doc)

  wrap_element(node)
}


# ============================================================================
# Child element descriptor helpers
# ============================================================================

#' Define a ZeroOrOne child element specification
#' @param nsptagname Namespace-prefixed tag, e.g. `"p:sldIdLst"`.
#' @param successors Character vector of successor tag names for ordering.
#' @return A list describing the child element spec.
#' @noRd
#' @export
zero_or_one <- function(nsptagname, successors = character(0)) {
  list(type = "zero_or_one", nsptagname = nsptagname, successors = successors)
}

#' Define a ZeroOrMore child element specification
#' @inheritParams zero_or_one
#' @return A list describing the child element spec.
#' @noRd
#' @export
zero_or_more <- function(nsptagname, successors = character(0)) {
  list(type = "zero_or_more", nsptagname = nsptagname, successors = successors)
}

#' Define a OneOrMore child element specification
#' @inheritParams zero_or_one
#' @return A list describing the child element spec.
#' @noRd
#' @export
one_or_more <- function(nsptagname, successors = character(0)) {
  list(type = "one_or_more", nsptagname = nsptagname, successors = successors)
}

#' Define a OneAndOnlyOne child element specification
#' @param nsptagname Namespace-prefixed tag.
#' @return A list describing the child element spec.
#' @noRd
#' @export
one_and_only_one <- function(nsptagname) {
  list(type = "one_and_only_one", nsptagname = nsptagname, successors = character(0))
}

#' Define an OptionalAttribute specification
#' @param attr_name Attribute name (may be namespace-prefixed).
#' @param simple_type A list with `from_xml` and `to_xml` functions.
#' @param default Default value when attribute is absent (default NULL).
#' @return A list describing the attribute spec.
#' @noRd
#' @export
optional_attribute <- function(attr_name, simple_type, default = NULL) {
  list(
    type = "optional_attribute",
    attr_name = attr_name,
    simple_type = simple_type,
    default = default
  )
}

#' Define a RequiredAttribute specification
#' @param attr_name Attribute name (may be namespace-prefixed).
#' @param simple_type A list with `from_xml` and `to_xml` functions.
#' @return A list describing the attribute spec.
#' @noRd
#' @export
required_attribute <- function(attr_name, simple_type) {
  list(
    type = "required_attribute",
    attr_name = attr_name,
    simple_type = simple_type
  )
}


# ============================================================================
# define_oxml_element() — the R equivalent of MetaOxmlElement
# ============================================================================

#' Define a custom XML element R6 class
#'
#' Factory function that generates an R6 class with active bindings for child
#' elements and attributes. This is the R equivalent of python-pptx's
#' MetaOxmlElement metaclass + xmlchemy descriptors.
#'
#' @param classname Name for the R6 class.
#' @param tag Namespace-prefixed tag, e.g. `"p:presentation"`.
#' @param children Named list of child element specs created with `zero_or_one()`,
#'   `zero_or_more()`, `one_or_more()`, or `one_and_only_one()`.
#' @param attributes Named list of attribute specs created with
#'   `optional_attribute()` or `required_attribute()`.
#' @param methods Named list of additional public methods (as functions).
#' @param active Named list of additional active bindings.
#' @param inherit R6 class to inherit from (default: BaseOxmlElement).
#' @return An R6ClassGenerator.
#' @noRd
#' @export
define_oxml_element <- function(classname,
                                 tag,
                                 children = list(),
                                 attributes = list(),
                                 methods = list(),
                                 active = list(),
                                 inherit = BaseOxmlElement) {

  # We build up public_methods, active_bindings from the child/attribute specs.
  # Because R6 replaces active binding function environments, we use bquote()
  # to bake literal values into function bodies.

  public_methods <- methods
  active_bindings <- active

  # --- Process child element specs ---
  for (prop_name in names(children)) {
    spec <- children[[prop_name]]
    child_type <- spec$type
    nsptagname <- spec$nsptagname
    successors <- spec$successors

    if (child_type == "zero_or_one") {
      # Active binding: getter returns child or NULL
      active_bindings[[prop_name]] <- .make_child_getter(nsptagname)

      # get_or_add_x(): return existing child or create + insert new one
      public_methods[[paste0("get_or_add_", prop_name)]] <-
        .make_get_or_add(prop_name, nsptagname, successors)

      # _add_x(): unconditionally create + insert a new child
      public_methods[[paste0("_add_", prop_name)]] <-
        .make_adder(prop_name, nsptagname, successors)

      # _new_x(): create a new loose element
      public_methods[[paste0("_new_", prop_name)]] <-
        .make_creator(nsptagname)

      # _insert_x(): insert element in the correct sequence
      public_methods[[paste0("_insert_", prop_name)]] <-
        .make_inserter(successors)

      # _remove_x(): remove all matching children
      public_methods[[paste0("_remove_", prop_name)]] <-
        .make_remover(nsptagname)

    } else if (child_type == "zero_or_more" || child_type == "one_or_more") {
      # List getter: x_lst active binding returns list of matching children
      active_bindings[[paste0(prop_name, "_lst")]] <-
        .make_list_getter(nsptagname)

      # _add_x(): unconditionally create + insert
      public_methods[[paste0("_add_", prop_name)]] <-
        .make_adder(prop_name, nsptagname, successors)

      # _new_x(): create new element
      public_methods[[paste0("_new_", prop_name)]] <-
        .make_creator(nsptagname)

      # _insert_x(): insert in sequence
      public_methods[[paste0("_insert_", prop_name)]] <-
        .make_inserter(successors)

      if (child_type == "one_or_more") {
        # add_x(): public adder
        public_methods[[paste0("add_", prop_name)]] <-
          .make_public_adder(prop_name)
      }

    } else if (child_type == "one_and_only_one") {
      # Active binding: getter returns child or raises error
      active_bindings[[prop_name]] <- .make_required_child_getter(nsptagname)
    }
  }

  # --- Process attribute specs ---
  for (prop_name in names(attributes)) {
    spec <- attributes[[prop_name]]
    attr_type <- spec$type
    attr_name <- spec$attr_name
    simple_type <- spec$simple_type

    # Resolve namespace-prefixed attribute names
    clark_attr <- if (grepl(":", attr_name, fixed = TRUE)) qn(attr_name) else attr_name

    if (attr_type == "optional_attribute") {
      default_val <- spec$default
      active_bindings[[prop_name]] <-
        .make_optional_attr_binding(clark_attr, simple_type, default_val)

    } else if (attr_type == "required_attribute") {
      active_bindings[[prop_name]] <-
        .make_required_attr_binding(clark_attr, attr_name, simple_type)
    }
  }

  # Build the R6 class
  R6::R6Class(
    classname = classname,
    inherit = inherit,
    public = public_methods,
    active = active_bindings
  )
}


# ============================================================================
# Internal factory functions for methods and active bindings
# ============================================================================
# All functions use bquote() to embed literal values directly into function
# bodies, avoiding closure issues when R6 replaces environments.

# --- Child element getters ---

.make_child_getter <- function(nsptagname) {
  # Returns an active binding function: getter returns child or NULL (read-only)
  eval(bquote(function(value) {
    if (!missing(value)) stop("This property is read-only.", call. = FALSE)
    self$find(qn(.(nsptagname)))
  }))
}

.make_required_child_getter <- function(nsptagname) {
  eval(bquote(function(value) {
    if (!missing(value)) stop("This property is read-only.", call. = FALSE)
    child <- self$find(qn(.(nsptagname)))
    if (is.null(child)) {
      invalid_xml_error(sprintf("required '<%s>' child element not present", .(nsptagname)))
    }
    child
  }))
}

.make_list_getter <- function(nsptagname) {
  eval(bquote(function(value) {
    if (!missing(value)) stop("This property is read-only.", call. = FALSE)
    self$findall(qn(.(nsptagname)))
  }))
}

# --- Child element mutation methods ---

.make_creator <- function(nsptagname) {
  eval(bquote(function() {
    OxmlElement(.(nsptagname))
  }))
}

.make_inserter <- function(successors) {
  eval(bquote(function(child) {
    self$insert_element_before(child, .(successors))
    child
  }))
}

.make_adder <- function(prop_name, nsptagname, successors) {
  new_method_name <- paste0("_new_", prop_name)
  insert_method_name <- paste0("_insert_", prop_name)
  eval(bquote(function(...) {
    new_method <- self[[.(new_method_name)]]
    child <- new_method()
    # Insert first so child is in the parent doc's namespace context
    insert_method <- self[[.(insert_method_name)]]
    insert_method(child)
    # Set attributes after insertion (namespace prefixes now available)
    attrs <- list(...)
    for (key in names(attrs)) {
      child[[key]] <- attrs[[key]]
    }
    child
  }))
}

.make_get_or_add <- function(prop_name, nsptagname, successors) {
  add_method_name <- paste0("_add_", prop_name)
  eval(bquote(function() {
    child <- self[[.(prop_name)]]
    if (is.null(child)) {
      add_method <- self[[.(add_method_name)]]
      child <- add_method()
    }
    child
  }))
}

.make_remover <- function(nsptagname) {
  eval(bquote(function() {
    self$remove_all(.(nsptagname))
  }))
}

.make_public_adder <- function(prop_name) {
  add_method_name <- paste0("_add_", prop_name)
  eval(bquote(function() {
    add_method <- self[[.(add_method_name)]]
    add_method()
  }))
}

# --- Attribute bindings ---

.make_optional_attr_binding <- function(clark_attr, simple_type, default_val) {
  # We need to embed simple_type (a list with from_xml/to_xml) into the function.
  # Use the environment trick: store in the function's own body via bquote.
  from_xml_fn <- simple_type$from_xml
  to_xml_fn <- simple_type$to_xml
  eval(bquote(function(value) {
    if (missing(value)) {
      # Getter
      attr_str <- self$get_attr(.(clark_attr))
      if (is.null(attr_str)) return(.(default_val))
      return((.(from_xml_fn))(attr_str))
    }
    # Setter
    if (identical(value, .(default_val))) {
      self$remove_attr(.(clark_attr))
    } else {
      str_value <- (.(to_xml_fn))(value)
      self$set_attr(.(clark_attr), str_value)
    }
  }))
}

.make_required_attr_binding <- function(clark_attr, attr_name, simple_type) {
  from_xml_fn <- simple_type$from_xml
  to_xml_fn <- simple_type$to_xml
  eval(bquote(function(value) {
    if (missing(value)) {
      # Getter
      attr_str <- self$get_attr(.(clark_attr))
      if (is.null(attr_str)) {
        invalid_xml_error(sprintf(
          "required '%s' attribute not present on element", .(attr_name)
        ))
      }
      return((.(from_xml_fn))(attr_str))
    }
    # Setter
    str_value <- (.(to_xml_fn))(value)
    self$set_attr(.(clark_attr), str_value)
  }))
}
