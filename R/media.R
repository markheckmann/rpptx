# Video — immutable value object representing a video file.
#
# Ported from python-pptx/src/pptx/media.py.

# MIME type → canonical file extension map
.video_ext_map <- c(
  "video/x-ms-asf"        = "asf",
  "video/avi"             = "avi",
  "video/quicktime"       = "mov",
  "video/mp4"             = "mp4",
  "video/mpeg"            = "mpg",
  "video/msvideo"         = "avi",
  "video/x-msvideo"       = "avi",
  "application/x-shockwave-flash" = "swf",
  "video/x-ms-wmv"        = "wmv"
)

#' Video value object
#'
#' Wraps a video bytestream and its MIME type. Used by `SlideShapes$add_movie()`.
#'
#' @importFrom digest digest
#' @keywords internal
#' @export
Video <- R6::R6Class(
  "Video",

  public = list(
    initialize = function(blob, mime_type = NULL, filename = NULL) {
      private$.blob      <- blob
      private$.mime_type <- mime_type %||% "video/unknown"
      private$.filename  <- filename
    }
  ),

  active = list(
    # Raw bytes of the video file.
    blob = function() private$.blob,

    # MIME type string, e.g. "video/mp4".
    content_type = function() private$.mime_type,

    # File extension (without dot), e.g. "mp4".
    ext = function() {
      if (!is.null(private$.filename)) {
        ext <- tools::file_ext(private$.filename)
        if (nchar(ext) > 0L) return(tolower(ext))
      }
      mapped <- unname(.video_ext_map[private$.mime_type])
      if (is.na(mapped)) "vid" else mapped
    },

    # Filename for the shape name, e.g. "video.mp4".
    filename = function() {
      if (!is.null(private$.filename)) return(private$.filename)
      sprintf("media.%s", self$ext)
    },

    # SHA1 hex digest of the blob for deduplication.
    sha1 = function() {
      digest::digest(private$.blob, algo = "sha1", serialize = FALSE)
    }
  ),

  private = list(.blob = NULL, .mime_type = NULL, .filename = NULL)
)

# Constructor: load video from a file path.
Video_from_file <- function(path, mime_type = NULL) {
  blob     <- readBin(path, "raw", n = file.info(path)$size)
  filename <- basename(path)
  Video$new(blob, mime_type, filename)
}

# Default "media loudspeaker" poster frame image (PNG).
.speaker_image_path <- function() {
  system.file("media-speaker.png", package = "rpptx")
}
