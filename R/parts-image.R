# ImagePart and Image — OPC part for image media files.
#
# Ported from python-pptx/src/pptx/parts/image.py.

#' @include opc-package.R
#' @noRd


# ============================================================================
# Image — immutable value object representing an image file
# ============================================================================

# Map of magic bytes → canonical extension.
.image_magic <- list(
  list(magic = as.raw(c(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a)), ext = "png",
       mime = "image/png"),
  list(magic = as.raw(c(0xff, 0xd8, 0xff)),                                 ext = "jpg",
       mime = "image/jpeg"),
  list(magic = as.raw(c(0x47, 0x49, 0x46, 0x38)),                           ext = "gif",
       mime = "image/gif"),  # "GIF8"
  list(magic = as.raw(c(0x42, 0x4d)),                                       ext = "bmp",
       mime = "image/bmp"),
  list(magic = as.raw(c(0x49, 0x49, 0x2a, 0x00)),                           ext = "tiff",
       mime = "image/tiff"),  # Little-endian TIFF
  list(magic = as.raw(c(0x4d, 0x4d, 0x00, 0x2a)),                           ext = "tiff",
       mime = "image/tiff")   # Big-endian TIFF
)

# Content-type map for recognised extensions.
.image_content_types <- c(
  png  = "image/png",
  jpg  = "image/jpeg",
  jpeg = "image/jpeg",
  gif  = "image/gif",
  bmp  = "image/bmp",
  tiff = "image/tiff",
  tif  = "image/tiff"
)


#' Image value object
#'
#' Immutable wrapper around raw image bytes that exposes ext, content_type,
#' sha1, size (pixels) and dpi.
#'
#' @noRd
Image <- R6::R6Class(
  "Image",

  public = list(
    initialize = function(blob, filename = NULL) {
      private$.blob     <- blob
      private$.filename <- filename
    },

    # MD5 hex digest of the image bytes (used for deduplication).
    sha1 = function() {
      tmp <- tempfile()
      writeBin(private$.blob, tmp)
      on.exit(unlink(tmp))
      unname(tools::md5sum(tmp))
    }
  ),

  active = list(
    blob          = function() private$.blob,
    filename      = function() private$.filename,

    # Canonical file extension e.g. "png".
    ext = function() {
      bytes <- private$.blob[seq_len(min(8L, length(private$.blob)))]
      for (entry in .image_magic) {
        n  <- length(entry$magic)
        if (length(bytes) >= n && identical(bytes[seq_len(n)], entry$magic)) {
          return(entry$ext)
        }
      }
      # Fallback: try filename extension
      if (!is.null(private$.filename)) {
        fe <- tolower(tools::file_ext(private$.filename))
        if (fe %in% names(.image_content_types)) return(fe)
      }
      stop("Unsupported or unrecognised image format", call. = FALSE)
    },

    # MIME type e.g. "image/png".
    content_type = function() {
      ct <- .image_content_types[self$ext]
      if (is.na(ct)) stop(sprintf("No content type for extension '%s'", self$ext), call. = FALSE)
      unname(ct)
    },

    # Pixel dimensions as integer vector c(width, height).
    size = function() private$.px_size(),

    # DPI as integer vector c(horz, vert).
    dpi = function() private$.dpi_value()
  ),

  private = list(
    .blob     = NULL,
    .filename = NULL,
    .props    = NULL,

    # Parse image properties using magick (if installed) or built-in readers.
    .get_props = function() {
      if (!is.null(private$.props)) return(private$.props)
      blob <- private$.blob
      ext  <- self$ext

      if (ext == "png") {
        props <- .png_props(blob)
      } else if (ext %in% c("jpg", "jpeg")) {
        props <- .jpeg_props(blob)
      } else if (requireNamespace("magick", quietly = TRUE)) {
        img   <- magick::image_read(blob)
        info  <- magick::image_info(img)
        props <- list(
          width  = info$width,
          height = info$height,
          x_dpi  = if (!is.na(info$density) && nchar(info$density) > 0) {
            as.integer(strsplit(info$density, "x")[[1]][1])
          } else 72L,
          y_dpi  = if (!is.na(info$density) && nchar(info$density) > 0) {
            as.integer(strsplit(info$density, "x")[[1]][2])
          } else 72L
        )
      } else {
        stop(sprintf(
          "Cannot read image dimensions for '%s' without the magick package", ext
        ), call. = FALSE)
      }
      private$.props <- props
      props
    },

    .px_size = function() {
      p <- private$.get_props()
      c(as.integer(p$width), as.integer(p$height))
    },

    .dpi_value = function() {
      p <- private$.get_props()
      horz <- as.integer(p$x_dpi)
      vert <- as.integer(p$y_dpi)
      if (is.na(horz) || horz < 1L) horz <- 72L
      if (is.na(vert) || vert < 1L) vert <- 72L
      c(horz, vert)
    }
  )
)


# Load an Image from a file path or raw connection.
Image_from_file <- function(image_file) {
  if (is.character(image_file)) {
    blob     <- readBin(image_file, "raw", file.info(image_file)$size)
    filename <- basename(image_file)
  } else {
    if (isSeekable(image_file)) seek(image_file, 0)
    blob     <- readBin(image_file, "raw", 1e7)
    filename <- NULL
  }
  Image$new(blob, filename)
}


# ============================================================================
# PNG / JPEG header readers (no external dependency)
# ============================================================================

# Read width, height, x_dpi, y_dpi from raw PNG bytes.
.png_props <- function(blob) {
  # Width at bytes 17-20, height at 21-24 (big-endian)
  w <- .be_int(blob[17:20])
  h <- .be_int(blob[21:24])
  # Look for pHYs chunk (physical pixel dimensions)
  x_dpi <- 72L; y_dpi <- 72L
  # Scan for "pHYs" signature (0x70485973)
  sig  <- as.raw(c(0x70, 0x48, 0x59, 0x73))
  for (i in seq_len(length(blob) - 16L)) {
    if (identical(blob[i:(i + 3L)], sig)) {
      px  <- .be_int(blob[(i + 4L):(i + 7L)])
      py  <- .be_int(blob[(i + 8L):(i + 11L)])
      unit <- as.integer(blob[i + 12L])
      if (unit == 1L) {   # metres^-1 → DPI
        x_dpi <- max(1L, as.integer(round(px / 39.3701)))
        y_dpi <- max(1L, as.integer(round(py / 39.3701)))
      }
      break
    }
  }
  list(width = w, height = h, x_dpi = x_dpi, y_dpi = y_dpi)
}

# Read width, height, x_dpi, y_dpi from raw JPEG bytes (APP0/APP1).
.jpeg_props <- function(blob) {
  # Default
  w <- 0L; h <- 0L; x_dpi <- 72L; y_dpi <- 72L

  i <- 3L  # skip FF D8 FF
  while (i < length(blob) - 4L) {
    if (!identical(blob[i], as.raw(0xff))) break
    marker <- blob[i + 1L]
    seg_len <- .be_int(blob[(i + 2L):(i + 3L)])  # includes 2-byte length field

    # APP0 (JFIF)
    if (identical(marker, as.raw(0xe0)) && seg_len >= 16L) {
      # Check for JFIF identifier "JFIF\0"
      jfif_sig <- as.raw(c(0x4a, 0x46, 0x49, 0x46, 0x00))
      if (i + 8L <= length(blob) && identical(blob[(i + 4L):(i + 8L)], jfif_sig)) {
        unit <- as.integer(blob[i + 11L])
        xd   <- .be_int(blob[(i + 12L):(i + 13L)])
        yd   <- .be_int(blob[(i + 14L):(i + 15L)])
        if (unit == 1L && xd > 0L && yd > 0L) { x_dpi <- xd; y_dpi <- yd }
      }
    }

    # SOF0 / SOF2 (0xC0, 0xC2) — contains width and height
    if (identical(marker, as.raw(0xc0)) || identical(marker, as.raw(0xc2))) {
      h <- .be_int(blob[(i + 5L):(i + 6L)])
      w <- .be_int(blob[(i + 7L):(i + 8L)])
      break
    }

    i <- i + 2L + seg_len
  }
  list(width = w, height = h, x_dpi = x_dpi, y_dpi = y_dpi)
}

# Read a big-endian integer from 2 or 4 raw bytes.
.be_int <- function(bytes) {
  n <- length(bytes)
  v <- 0L
  for (k in seq_len(n)) {
    v <- bitwOr(bitwShiftL(v, 8L), as.integer(bytes[k]))
  }
  v
}


# ============================================================================
# ImagePart — OPC Part for image files
# ============================================================================

#' Image OPC part
#'
#' Stores the raw bytes of an image and provides dimension scaling.
#'
#' @noRd
ImagePart <- R6::R6Class(
  "ImagePart",
  inherit = Part,

  public = list(
    initialize = function(partname, content_type, package, blob, filename = NULL) {
      super$initialize(partname, content_type, package, blob)
      private$.blob     <- blob
      private$.filename <- filename
    },

    # Return (cx, cy) in EMU, filling in any NULL dimension by aspect-ratio.
    scale = function(cx, cy) {
      img_size <- private$.native_size()
      img_cx   <- img_size[1]; img_cy <- img_size[2]

      if (!is.null(cx) && !is.null(cy)) return(c(as.integer(cx), as.integer(cy)))

      if (!is.null(cx)) {
        factor <- cx / img_cx
        return(c(as.integer(cx), as.integer(round(img_cy * factor))))
      }
      if (!is.null(cy)) {
        factor <- cy / img_cy
        return(c(as.integer(round(img_cx * factor)), as.integer(cy)))
      }
      c(as.integer(img_cx), as.integer(img_cy))
    }
  ),

  active = list(
    # Filename for use in `descr` attribute of the p:pic element.
    desc = function() {
      if (!is.null(private$.filename)) return(private$.filename)
      paste0("image.", tools::file_ext(self$partname))
    },

    sha1 = function() {
      tmp <- tempfile()
      writeBin(private$.blob, tmp)
      on.exit(unlink(tmp))
      unname(tools::md5sum(tmp))
    }
  ),

  private = list(
    .blob     = NULL,
    .filename = NULL,

    .native_size = function() {
      img <- Image$new(private$.blob, private$.filename)
      px  <- img$size
      dpi <- img$dpi
      EMU_PER_INCH <- 914400L
      cx  <- as.integer(round(EMU_PER_INCH * px[1] / dpi[1]))
      cy  <- as.integer(round(EMU_PER_INCH * px[2] / dpi[2]))
      c(cx, cy)
    }
  )
)


# Create a new ImagePart from an Image value object.
ImagePart_new <- function(package, image) {
  partname <- package$next_image_partname(image$ext)
  ImagePart$new(partname, image$content_type, package, image$blob, image$filename)
}
