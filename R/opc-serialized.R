# API for reading/writing serialized OPC packages.
#
# Ported from python-pptx/src/pptx/opc/serialized.py. Provides PackageReader
# and PackageWriter for ZIP-format OPC packages.

# ============================================================================
# PackageReader — dict-like access to package parts
# ============================================================================

#' Read an OPC package (ZIP file)
#'
#' Provides dict-like access to the binary blobs of each part in the package.
#'
#' @include opc-packuri.R opc-oxml.R opc-constants.R opc-spec.R
#' @noRd
#' @export
PackageReader <- R6::R6Class(
  "PackageReader",

  public = list(
    initialize = function(pkg_file) {
      private$.pkg_file <- pkg_file
    },

    # Check if a part exists in the package
    # @param pack_uri A PackURI.
    # @return Logical.
    contains = function(pack_uri) {
      as.character(pack_uri) %in% names(private$.blobs())
    },

    # Get the blob (raw bytes) for a part
    # @param pack_uri A PackURI.
    # @return A raw vector.
    get_blob = function(pack_uri) {
      blobs <- private$.blobs()
      key <- as.character(pack_uri)
      if (!(key %in% names(blobs))) {
        stop(sprintf("no member '%s' in package", key), call. = FALSE)
      }
      blobs[[key]]
    },

    # Get the rels XML bytes for a part (or NULL if no .rels exists)
    # @param partname A PackURI.
    # @return Raw bytes or NULL.
    rels_xml_for = function(partname) {
      rels_uri <- pack_uri_rels_uri(partname)
      if (self$contains(rels_uri)) {
        return(self$get_blob(rels_uri))
      }
      NULL
    }
  ),

  private = list(
    .pkg_file = NULL,
    .blobs_cache = NULL,

    .blobs = function() {
      if (is.null(private$.blobs_cache)) {
        private$.blobs_cache <- .read_zip_blobs(private$.pkg_file)
      }
      private$.blobs_cache
    }
  )
)


# ============================================================================
# PackageWriter — write parts to a ZIP package
# ============================================================================

#' Write an OPC package (ZIP file)
#'
#' @noRd
#' @export
PackageWriter <- R6::R6Class(
  "PackageWriter",

  public = list(
    # Write a physical package (.pptx file)
    #
    # @param pkg_file Path to output file.
    # @param pkg_rels The package-level Relationships object.
    # @param parts A list of Part objects to write.
    write = function(pkg_file, pkg_rels, parts) {
      # Collect all blobs: list of (membername, raw_bytes)
      items <- list()

      # 1. Content types
      ct_xml <- .content_types_xml_for(parts)
      items[[length(items) + 1]] <- list(
        membername = pack_uri_membername(CONTENT_TYPES_URI),
        blob = ct_xml
      )

      # 2. Package rels
      pkg_rels_xml <- pkg_rels$xml_bytes()
      items[[length(items) + 1]] <- list(
        membername = pack_uri_membername(pack_uri_rels_uri(PACKAGE_URI)),
        blob = pkg_rels_xml
      )

      # 3. Parts and their rels
      for (part in parts) {
        items[[length(items) + 1]] <- list(
          membername = pack_uri_membername(part$partname),
          blob = part$blob
        )
        if (part$has_rels()) {
          rels_xml <- part$rels_xml_bytes()
          items[[length(items) + 1]] <- list(
            membername = pack_uri_membername(pack_uri_rels_uri(part$partname)),
            blob = rels_xml
          )
        }
      }

      # Write all items to a ZIP file
      .write_zip_package(pkg_file, items)
    }
  )
)


# ============================================================================
# Internal ZIP I/O helpers
# ============================================================================

#' Read all blobs from a ZIP file
#' @param pkg_file Path to a .pptx file or a raw connection.
#' @return Named list mapping PackURI strings to raw vectors.
#' @noRd
.read_zip_blobs <- function(pkg_file) {
  if (is.character(pkg_file)) {
    if (!file.exists(pkg_file)) {
      package_not_found_error(pkg_file)
    }
  }

  # Use zip package to list and read entries
  tmpdir <- tempfile("rpptx_")
  dir.create(tmpdir, recursive = TRUE)
  on.exit(unlink(tmpdir, recursive = TRUE), add = TRUE)

  utils::unzip(pkg_file, exdir = tmpdir)

  # Walk the extracted directory and read all files
  files <- list.files(tmpdir, recursive = TRUE, all.files = TRUE)
  blobs <- list()
  for (f in files) {
    full_path <- file.path(tmpdir, f)
    blob <- readBin(full_path, what = "raw", n = file.info(full_path)$size)
    pack_uri <- PackURI(paste0("/", f))
    blobs[[as.character(pack_uri)]] <- blob
  }
  blobs
}


#' Write items to a ZIP file
#' @param pkg_file Path to output file.
#' @param items A list of lists with `membername` and `blob` elements.
#' @noRd
.write_zip_package <- function(pkg_file, items) {
  tmpdir <- tempfile("rpptx_write_")
  dir.create(tmpdir, recursive = TRUE)
  on.exit(unlink(tmpdir, recursive = TRUE), add = TRUE)

  # Write each blob to a file in the temp directory
  membernames <- character(length(items))
  for (i in seq_along(items)) {
    membername <- items[[i]]$membername
    blob <- items[[i]]$blob
    membernames[i] <- membername

    full_path <- file.path(tmpdir, membername)
    dir.create(dirname(full_path), recursive = TRUE, showWarnings = FALSE)
    writeBin(blob, full_path)
  }

  # Create ZIP — use utils::zip which preserves directory structure
  old_wd <- getwd()
  setwd(tmpdir)
  on.exit(setwd(old_wd), add = TRUE)

  utils::zip(
    normalizePath(pkg_file, mustWork = FALSE),
    files = membernames,
    flags = "-q"
  )
}


# ============================================================================
# Content types generation
# ============================================================================

#' Generate \[Content_Types\].xml bytes for a set of parts
#' @param parts A list of Part objects.
#' @return Raw bytes of the content types XML.
#' @noRd
.content_types_xml_for <- function(parts) {
  # Separate into defaults (by extension) and overrides (by partname)
  defaults <- list(rels = CT$OPC_RELATIONSHIPS, xml = CT$XML)
  overrides <- list()

  for (part in parts) {
    partname <- part$partname
    content_type <- part$content_type
    ext <- tolower(pack_uri_ext(partname))

    if (is_default_content_type(ext, content_type)) {
      defaults[[ext]] <- content_type
    } else {
      overrides[[as.character(partname)]] <- content_type
    }
  }

  # Build the CT_Types element
  ct_types <- new_ct_types()

  # Add defaults sorted by extension
  for (ext in sort(names(defaults))) {
    ct_types$add_default(ext, defaults[[ext]])
  }
  # Add overrides sorted by partname
  for (partname in sort(names(overrides))) {
    ct_types$add_override(partname, overrides[[partname]])
  }

  serialize_part_xml(ct_types)
}


# ============================================================================
# CaseInsensitiveDict helper
# ============================================================================

#' A simple case-insensitive named list
#' @param ... Key-value pairs.
#' @return An environment with case-insensitive access.
#' @noRd
CaseInsensitiveDict <- function(...) {
  items <- list(...)
  result <- list()
  for (nm in names(items)) {
    result[[tolower(nm)]] <- items[[nm]]
  }
  result
}
