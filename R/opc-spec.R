# Mappings that embody aspects of the Open XML spec ISO/IEC 29500.
#
# Ported from python-pptx/src/pptx/opc/spec.py.

# Default content types: pairs of (ext, content_type) where the content type
# can be inferred from the file extension alone.
# Uses CT (CONTENT_TYPE) constants defined in opc-constants.R.

# Use a local environment for mutable state (avoids locked binding issues)
.spec_env <- new.env(parent = emptyenv())
.spec_env$default_content_types <- NULL

.init_default_content_types <- function() {
  if (is.null(.spec_env$default_content_types)) {
    .spec_env$default_content_types <- list(
      c("bin", CT$PML_PRINTER_SETTINGS),
      c("bin", CT$SML_PRINTER_SETTINGS),
      c("bin", CT$WML_PRINTER_SETTINGS),
      c("bmp", CT$BMP),
      c("emf", CT$X_EMF),
      c("fntdata", CT$X_FONTDATA),
      c("gif", CT$GIF),
      c("jpe", CT$JPEG),
      c("jpeg", CT$JPEG),
      c("jpg", CT$JPEG),
      c("mov", CT$MOV),
      c("mp4", CT$MP4),
      c("mpg", CT$MPG),
      c("png", CT$PNG),
      c("rels", CT$OPC_RELATIONSHIPS),
      c("tif", CT$TIFF),
      c("tiff", CT$TIFF),
      c("vid", CT$VIDEO),
      c("wdp", CT$MS_PHOTO),
      c("wmf", CT$X_WMF),
      c("wmv", CT$WMV),
      c("xlsx", CT$SML_SHEET),
      c("xml", CT$XML)
    )
  }
  .spec_env$default_content_types
}

#' Check if a (ext, content_type) pair is a default content type
#' @param ext File extension (lowercase).
#' @param content_type Content type string.
#' @return Logical.
#' @keywords internal
is_default_content_type <- function(ext, content_type) {
  dct <- .init_default_content_types()
  pair <- c(tolower(ext), content_type)
  any(vapply(dct, function(x) identical(x, pair), logical(1)))
}

#' @keywords internal
image_content_types <- list(
  bmp  = "image/bmp",
  emf  = "image/x-emf",
  gif  = "image/gif",
  jpe  = "image/jpeg",
  jpeg = "image/jpeg",
  jpg  = "image/jpeg",
  png  = "image/png",
  tif  = "image/tiff",
  tiff = "image/tiff",
  wdp  = "image/vnd.ms-photo",
  wmf  = "image/x-wmf"
)
