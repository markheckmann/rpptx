# PackURI value type and known pack-URI strings.
#
# Ported from python-pptx/src/pptx/opc/packuri.py. PackURI is an S3 class
# wrapping a string that provides utility properties for manipulating OPC
# part URIs.

#' Create a PackURI (OPC part name)
#'
#' A PackURI is a string that represents an absolute URI within an OPC package.
#' It must begin with a forward slash.
#'
#' @param uri_str A string starting with "/", e.g. `"/ppt/slides/slide1.xml"`.
#' @return A PackURI (character with S3 class).
#' @noRd
#' @export
PackURI <- function(uri_str) {
  if (!is.character(uri_str) || length(uri_str) != 1) {
    stop("PackURI must be a single character string", call. = FALSE)
  }
  if (!startsWith(uri_str, "/")) {
    stop(sprintf("PackURI must begin with slash, got '%s'", uri_str), call. = FALSE)
  }
  structure(uri_str, class = c("PackURI", "character"))
}

#' Construct a PackURI from a relative reference
#'
#' @param base_uri The base URI (directory) for resolving the reference.
#' @param relative_ref The relative reference string.
#' @return A PackURI.
#' @noRd
#' @export
pack_uri_from_rel_ref <- function(base_uri, relative_ref) {
  # Use POSIX-style path joining and normalization
  joined <- paste0(base_uri, "/", relative_ref)
  abs_uri <- .posix_normpath(joined)
  PackURI(abs_uri)
}

#' @export
print.PackURI <- function(x, ...) {
  cat(sprintf("<PackURI: %s>\n", as.character(x)))
  invisible(x)
}

#' Get the base URI (directory portion) of a PackURI
#' @param uri A PackURI.
#' @return A character string.
#' @noRd
#' @export
pack_uri_base <- function(uri) {
  dirname(as.character(uri))
}

#' Get the file extension of a PackURI (without leading dot)
#' @param uri A PackURI.
#' @return A character string.
#' @noRd
#' @export
pack_uri_ext <- function(uri) {
  ext <- tools::file_ext(as.character(uri))
  ext
}

#' Get the filename portion of a PackURI
#' @param uri A PackURI.
#' @return A character string.
#' @noRd
#' @export
pack_uri_filename <- function(uri) {
  basename(as.character(uri))
}

#' Get the integer index from an array-style PackURI
#'
#' Returns the trailing integer from the filename, e.g. 21 for
#' `"/ppt/slides/slide21.xml"`, or NULL for singleton parts.
#'
#' @param uri A PackURI.
#' @return An integer or NULL.
#' @noRd
#' @export
pack_uri_idx <- function(uri) {
  filename <- pack_uri_filename(uri)
  if (filename == "" || filename == "/") return(NULL)
  name_part <- tools::file_path_sans_ext(filename)
  m <- regmatches(name_part, regexec("^[a-zA-Z]+([0-9]+)$", name_part))[[1]]
  if (length(m) < 2) return(NULL)
  as.integer(m[2])
}

#' Get the membername (without leading slash) for ZIP storage
#' @param uri A PackURI.
#' @return A character string.
#' @noRd
#' @export
pack_uri_membername <- function(uri) {
  sub("^/", "", as.character(uri))
}

#' Get a relative reference from a PackURI to a base URI
#' @param uri A PackURI.
#' @param base_uri The base URI to compute the relative reference from.
#' @return A character string.
#' @noRd
#' @export
pack_uri_relative_ref <- function(uri, base_uri) {
  uri_str <- as.character(uri)
  if (base_uri == "/") {
    return(sub("^/", "", uri_str))
  }
  # Compute POSIX relative path
  .posix_relpath(uri_str, base_uri)
}

#' Get the .rels URI corresponding to a PackURI
#' @param uri A PackURI.
#' @return A PackURI for the .rels part.
#' @noRd
#' @export
pack_uri_rels_uri <- function(uri) {
  filename <- pack_uri_filename(uri)
  rels_filename <- paste0(filename, ".rels")
  base <- pack_uri_base(uri)
  # Use .posix_normpath to handle double slashes when base is "/"
  rels_uri_str <- .posix_normpath(paste0(base, "/_rels/", rels_filename))
  PackURI(rels_uri_str)
}


# --- Well-known URIs ---

#' @noRd
PACKAGE_URI <- PackURI("/")

#' @noRd
CONTENT_TYPES_URI <- PackURI("/[Content_Types].xml")


# --- Internal POSIX path helpers ---

.posix_normpath <- function(path) {
  # Normalize a POSIX-style path (resolve ".." and ".")
  parts <- strsplit(path, "/", fixed = TRUE)[[1]]
  stack <- character(0)
  for (p in parts) {
    if (p == "" || p == ".") {
      next
    } else if (p == "..") {
      if (length(stack) > 0) stack <- stack[-length(stack)]
    } else {
      stack <- c(stack, p)
    }
  }
  result <- paste0("/", paste(stack, collapse = "/"))
  result
}

.posix_relpath <- function(path, start) {
  # Compute relative path from start to path (POSIX-style)
  path_parts <- strsplit(sub("^/", "", path), "/", fixed = TRUE)[[1]]
  start_parts <- strsplit(sub("^/", "", start), "/", fixed = TRUE)[[1]]

  # Find common prefix length
  common <- 0L
  max_common <- min(length(path_parts), length(start_parts))
  for (i in seq_len(max_common)) {
    if (path_parts[i] == start_parts[i]) {
      common <- i
    } else {
      break
    }
  }

  # Build relative path
  up_count <- length(start_parts) - common
  remaining <- path_parts[(common + 1):length(path_parts)]
  parts <- c(rep("..", up_count), remaining)
  paste(parts, collapse = "/")
}
