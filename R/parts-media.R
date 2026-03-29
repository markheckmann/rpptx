# MediaPart — OPC part for audio/video media.
#
# Ported from python-pptx/src/pptx/parts/media.py.

#' Media part (audio/video OPC part)
#'
#' @keywords internal
#' @export
MediaPart <- R6::R6Class(
  "MediaPart",
  inherit = Part,

  active = list(
    # SHA1 hex digest of the media blob for deduplication.
    sha1 = function() digest::digest(self$blob, algo = "sha1", serialize = FALSE)
  )
)

# Construct a new MediaPart from a Video object and register it in the package.
MediaPart_new <- function(package, video) {
  partname     <- package$next_media_partname(video$ext)
  content_type <- video$content_type
  blob         <- video$blob
  part         <- MediaPart$new(partname, content_type, package, blob)
  part
}
