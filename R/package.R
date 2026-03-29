# Package — top-level container for a .pptx file.
#
# Ported from python-pptx/src/pptx/package.py. Extends OpcPackage with
# PowerPoint-specific functionality.

#' PowerPoint package
#'
#' Top-level container for a .pptx file. Extends OpcPackage with
#' presentation-specific behavior.
#'
#' @include opc-package.R parts-coreprops.R parts-image.R
#' @noRd
Package <- R6::R6Class(
  "Package",
  inherit = OpcPackage,

  public = list(
    # Return the next available image partname with the given extension.
    next_image_partname = function(ext) {
      existing <- Filter(
        function(p) grepl("^/ppt/media/image[0-9]+\\.", p$partname),
        self$iter_parts()
      )
      idxs <- as.integer(regmatches(
        sapply(existing, function(p) p$partname),
        regexpr("[0-9]+(?=\\.)", sapply(existing, function(p) p$partname), perl = TRUE)
      ))
      idx <- 1L
      for (candidate in seq_len(length(idxs) + 1L)) {
        if (!(candidate %in% idxs)) { idx <- candidate; break }
      }
      PackURI(sprintf("/ppt/media/image%d.%s", idx, ext))
    },

    # Return the next available media partname (ppt/media/mediaN.ext).
    next_media_partname = function(ext) {
      existing <- Filter(
        function(p) grepl("^/ppt/media/media[0-9]+\\.", p$partname),
        self$iter_parts()
      )
      idxs <- if (length(existing) == 0L) integer(0L) else
        as.integer(regmatches(
          sapply(existing, function(p) p$partname),
          regexpr("[0-9]+(?=\\.)", sapply(existing, function(p) p$partname), perl = TRUE)
        ))
      idx <- 1L
      for (candidate in seq_len(length(idxs) + 1L)) {
        if (!(candidate %in% idxs)) { idx <- candidate; break }
      }
      PackURI(sprintf("/ppt/media/media%d.%s", idx, ext))
    },

    # Return a MediaPart for the given Video; reuse if same content already exists.
    get_or_add_media_part = function(video) {
      hash <- video$sha1
      for (p in self$iter_parts()) {
        if (inherits(p, "MediaPart")) {
          if (identical(p$sha1, hash)) return(p)
        }
      }
      MediaPart_new(self, video)
    },

    # Return an ImagePart for the given image file; reuse if same image already exists.
    get_or_add_image_part = function(image_file) {
      image <- Image_from_file(image_file)
      hash  <- image$sha1()
      # Look for existing image part with same content
      for (p in self$iter_parts()) {
        if (inherits(p, "ImagePart") && !is.null(tryCatch(p$sha1, error = function(e) NULL))) {
          if (identical(p$sha1, hash)) return(p)
        }
      }
      ImagePart_new(self, image)
    },

    # Iterate all parts reachable from this package.
    iter_parts = function() {
      seen  <- list()
      queue <- list(self)
      parts <- list()
      while (length(queue) > 0L) {
        node  <- queue[[1L]]
        queue <- queue[-1L]
        for (rel in node$rels$.__enclos_env__$private$.rels) {
          if (rel$is_external) next
          part <- rel$target_part
          key  <- part$partname
          if (key %in% seen) next
          seen  <- c(seen, key)
          parts <- c(parts, list(part))
          queue <- c(queue, list(part))
        }
      }
      parts
    }
  ),

  active = list(
    # The PresentationPart for this package
    presentation_part = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      self$main_document_part
    },

    # CorePropertiesPart (lazy; creates a default if absent)
    core_properties = function(value) {
      if (!missing(value)) stop("Read-only property", call. = FALSE)
      if (is.null(private$.core_properties)) {
        private$.core_properties <- tryCatch(
          self$part_related_by(RT$CORE_PROPERTIES),
          error = function(e) {
            cp <- CorePropertiesPart_default(self)
            self$relate_to(cp, RT$CORE_PROPERTIES)
            cp
          }
        )
      }
      private$.core_properties
    }
  ),

  private = list(
    .core_properties = NULL
  )
)

#' Open a Package from a .pptx file
#' @param pkg_file Path to a .pptx file.
#' @return A Package instance.
#' @noRd
Package_open <- function(pkg_file) {
  pkg <- Package$new(pkg_file)
  pkg$.__enclos_env__$private$.load()
  pkg
}
