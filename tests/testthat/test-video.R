# Tests for Video embedding (add_movie)

test_files <- system.file("test_files", package = "rpptx")
mp4_file   <- file.path(test_files, "dummy.mp4")
img_file   <- file.path(test_files, "monty-truth.png")

pptx_path <- function(name) system.file("test_files", name, package = "rpptx")

# Helper: fresh presentation + blank slide (no placeholder shapes)
new_blank_slide <- function() {
  prs    <- pptx_presentation(pptx_path("no-slides.pptx"))
  layout <- prs$slide_masters[[1]]$slide_layouts[[1]]
  slide  <- prs$slides$add_slide(layout)
  list(prs = prs, slide = slide)
}

# ============================================================================
# Video value object
# ============================================================================

describe("Video", {
  it("reads blob and reports ext from filename", {
    v <- Video_from_file(mp4_file)
    expect_true(is.raw(v$blob))
    expect_gt(length(v$blob), 0L)
    expect_equal(v$ext, "mp4")
    expect_equal(v$filename, "dummy.mp4")
  })

  it("derives ext from MIME type when filename lacks extension", {
    raw_data <- as.raw(c(0x00, 0x01))
    v <- Video$new(raw_data, "video/x-ms-wmv", NULL)
    expect_equal(v$ext, "wmv")
    expect_equal(v$filename, "media.wmv")
  })

  it("falls back to 'vid' for unknown MIME type", {
    raw_data <- as.raw(0x00)
    v <- Video$new(raw_data, "video/unknown", NULL)
    expect_equal(v$ext, "vid")
  })

  it("produces consistent SHA1 digest", {
    v1 <- Video_from_file(mp4_file)
    v2 <- Video_from_file(mp4_file)
    expect_equal(v1$sha1, v2$sha1)
    expect_type(v1$sha1, "character")
    expect_equal(nchar(v1$sha1), 40L)
  })
})

# ============================================================================
# add_movie
# ============================================================================

describe("SlideShapes$add_movie", {
  it("adds a picture shape to the slide", {
    s     <- new_blank_slide()
    n_before <- length(s$slide$shapes)

    shape <- s$slide$shapes$add_movie(mp4_file,
                                      Inches(1), Inches(1), Inches(4), Inches(3))
    expect_s3_class(shape, "R6")
    expect_equal(length(s$slide$shapes), n_before + 1L)
  })

  it("sets the shape position and size correctly", {
    s     <- new_blank_slide()

    shape <- s$slide$shapes$add_movie(mp4_file,
                                      Inches(2), Inches(3), Inches(5), Inches(4))
    expect_equal(shape$left,   Inches(2))
    expect_equal(shape$top,    Inches(3))
    expect_equal(shape$width,  Inches(5))
    expect_equal(shape$height, Inches(4))
  })

  it("creates media and video relationships on the slide part", {
    s <- new_blank_slide()
    s$slide$shapes$add_movie(mp4_file,
                             Inches(1), Inches(1), Inches(4), Inches(3))

    part <- s$prs$slides[[1]]$part
    rel_types <- vapply(
      part$rels$.__enclos_env__$private$.rels,
      function(r) r$reltype,
      character(1)
    )
    expect_true(RT$MEDIA %in% rel_types)
    expect_true(RT$VIDEO %in% rel_types)
    expect_true(RT$IMAGE %in% rel_types)
  })

  it("adds timing element with p:video to the slide XML", {
    s <- new_blank_slide()
    s$slide$shapes$add_movie(mp4_file,
                             Inches(1), Inches(1), Inches(4), Inches(3))

    sld_node <- s$slide$element$get_node()
    timing <- xml2::xml_find_first(sld_node, "p:timing",
                                   ns = c(p = .nsmap[["p"]]))
    expect_false(inherits(timing, "xml_missing"))

    video_el <- xml2::xml_find_first(sld_node,
                                     ".//p:video",
                                     ns = c(p = .nsmap[["p"]]))
    expect_false(inherits(video_el, "xml_missing"))
  })

  it("accepts a custom poster frame image", {
    s <- new_blank_slide()

    shape <- s$slide$shapes$add_movie(mp4_file,
                                      Inches(1), Inches(1), Inches(4), Inches(3),
                                      poster_frame_image = img_file)
    expect_s3_class(shape, "R6")
  })

  it("deduplicates media part for same video file", {
    s        <- new_blank_slide()
    n_before <- length(s$slide$shapes)

    s$slide$shapes$add_movie(mp4_file,
                             Inches(1), Inches(1), Inches(4), Inches(3))
    s$slide$shapes$add_movie(mp4_file,
                             Inches(6), Inches(1), Inches(4), Inches(3))

    media_parts <- Filter(
      function(p) inherits(p, "MediaPart"),
      s$slide$part$package$iter_parts()
    )
    expect_equal(length(media_parts), 1L)
    expect_equal(length(s$slide$shapes), n_before + 2L)
  })

  it("round-trips through save/reopen", {
    s <- new_blank_slide()
    n_before <- length(s$slide$shapes)
    s$slide$shapes$add_movie(mp4_file,
                             Inches(1), Inches(1), Inches(4), Inches(3))

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    s$prs$save(tmp)
    expect_true(file.exists(tmp))
    expect_gt(file.info(tmp)$size, 0L)

    prs2 <- pptx_presentation(tmp)
    expect_equal(length(prs2$slides), 1L)
    expect_equal(length(prs2$slides[[1]]$shapes), n_before + 1L)
  })
})
