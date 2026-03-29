# Tests for slide-related domain objects (Phase 3).

library(rpptx)

pptx_files <- list.files(
  system.file("test-pptx", package = "rpptx"),
  pattern = "\\.pptx$", full.names = TRUE
)

# ============================================================================
# SlideMasters / SlideLayouts collections
# ============================================================================

test_that("slide masters are accessible from default template", {
  prs <- pptx_presentation()
  masters <- prs$slide_masters
  expect_s3_class(masters, "SlideMasters")
  expect_equal(length(masters), 1L)
})

test_that("SlideMasters [[ returns SlideMaster", {
  prs <- pptx_presentation()
  master <- prs$slide_masters[[1]]
  expect_s3_class(master, "SlideMaster")
})

test_that("slide layouts are accessible from master", {
  prs <- pptx_presentation()
  master <- prs$slide_masters[[1]]
  layouts <- master$slide_layouts
  expect_s3_class(layouts, "SlideLayouts")
  expect_gt(length(layouts), 0L)
})

test_that("SlideLayouts [[ returns SlideLayout", {
  prs <- pptx_presentation()
  layout <- prs$slide_layouts[[1]]
  expect_s3_class(layout, "SlideLayout")
})

test_that("slide layout has a name", {
  prs <- pptx_presentation()
  layout <- prs$slide_layouts[[1]]
  expect_type(layout$name, "character")
  expect_gt(nchar(layout$name), 0L)
})

test_that("prs$slide_layouts returns layouts of first master", {
  prs <- pptx_presentation()
  n_via_prs <- length(prs$slide_layouts)
  n_via_master <- length(prs$slide_masters[[1]]$slide_layouts)
  expect_equal(n_via_prs, n_via_master)
})


# ============================================================================
# Slides — add_slide, indexing, slide_id
# ============================================================================

test_that("default presentation has no slides", {
  prs <- pptx_presentation()
  expect_equal(length(prs$slides), 0L)
})

test_that("add_slide returns a Slide", {
  prs <- pptx_presentation()
  layout <- prs$slide_layouts[[1]]
  slide <- prs$slides$add_slide(layout)
  expect_s3_class(slide, "Slide")
})

test_that("add_slide increments length", {
  prs <- pptx_presentation()
  layout <- prs$slide_layouts[[1]]
  prs$slides$add_slide(layout)
  expect_equal(length(prs$slides), 1L)
  prs$slides$add_slide(layout)
  expect_equal(length(prs$slides), 2L)
})

test_that("slide IDs are unique and within valid range", {
  prs <- pptx_presentation()
  layout <- prs$slide_layouts[[1]]
  s1 <- prs$slides$add_slide(layout)
  s2 <- prs$slides$add_slide(layout)
  expect_true(s1$slide_id != s2$slide_id)
  expect_gte(s1$slide_id, 256L)
  expect_lte(s1$slide_id, 2147483647L)
})

test_that("[[ indexing returns the correct slide", {
  prs <- pptx_presentation()
  layout <- prs$slide_layouts[[1]]
  s1 <- prs$slides$add_slide(layout)
  s2 <- prs$slides$add_slide(layout)
  expect_equal(prs$slides[[1]]$slide_id, s1$slide_id)
  expect_equal(prs$slides[[2]]$slide_id, s2$slide_id)
})

test_that("slide knows its layout", {
  prs <- pptx_presentation()
  layout <- prs$slide_layouts[[1]]
  slide <- prs$slides$add_slide(layout)
  expect_equal(slide$slide_layout$name, layout$name)
})


# ============================================================================
# Round-trip save/reload preserves slides
# ============================================================================

test_that("slides survive save/reopen round-trip", {
  prs <- pptx_presentation()
  layout1 <- prs$slide_layouts[[1]]
  layout2 <- prs$slide_layouts[[2]]
  prs$slides$add_slide(layout1)
  prs$slides$add_slide(layout2)

  tmp <- tempfile(fileext = ".pptx")
  on.exit(unlink(tmp))
  prs$save(tmp)

  prs2 <- pptx_presentation(tmp)
  expect_equal(length(prs2$slides), 2L)
  expect_equal(prs2$slides[[1]]$slide_layout$name, layout1$name)
  expect_equal(prs2$slides[[2]]$slide_layout$name, layout2$name)
})

test_that("slide_id is stable across round-trip", {
  prs <- pptx_presentation()
  layout <- prs$slide_layouts[[1]]
  slide <- prs$slides$add_slide(layout)
  original_id <- slide$slide_id

  tmp <- tempfile(fileext = ".pptx")
  on.exit(unlink(tmp))
  prs$save(tmp)

  prs2 <- pptx_presentation(tmp)
  expect_equal(prs2$slides[[1]]$slide_id, original_id)
})


# ============================================================================
# Slide deletion
# ============================================================================

describe("Slides$delete", {
  it("removes a slide from the collection", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    s1 <- prs$slides$add_slide(layout)
    s2 <- prs$slides$add_slide(layout)
    s3 <- prs$slides$add_slide(layout)
    prs$slides$delete(s2)
    expect_equal(length(prs$slides), 2L)
  })

  it("persists deletion across save/reload", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    s1 <- prs$slides$add_slide(layout)
    s2 <- prs$slides$add_slide(layout)
    prs$slides$delete(s1)

    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    expect_equal(length(prs2$slides), 1L)
  })

  it("errors when slide not in presentation", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    s1 <- prs$slides$add_slide(layout)

    prs2   <- pptx_presentation()
    layout2 <- prs2$slide_layouts[[1]]
    s2 <- prs2$slides$add_slide(layout2)

    expect_error(prs$slides$delete(s2), regexp = "not found")
  })
})


# ============================================================================
# Slide reordering
# ============================================================================

describe("Slides$move", {
  it("moves a slide forward in the collection", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    s1 <- prs$slides$add_slide(layout)
    s2 <- prs$slides$add_slide(layout)
    s3 <- prs$slides$add_slide(layout)
    ids_orig <- sapply(prs$slides$to_list(), function(s) s$slide_id)

    prs$slides$move(s1, 3L)
    ids_new <- sapply(prs$slides$to_list(), function(s) s$slide_id)

    # s1 (originally first) should now be last
    expect_equal(ids_new, c(ids_orig[2], ids_orig[3], ids_orig[1]))
  })

  it("moves a slide backward in the collection", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    s1 <- prs$slides$add_slide(layout)
    s2 <- prs$slides$add_slide(layout)
    s3 <- prs$slides$add_slide(layout)
    ids_orig <- sapply(prs$slides$to_list(), function(s) s$slide_id)

    prs$slides$move(s3, 1L)
    ids_new <- sapply(prs$slides$to_list(), function(s) s$slide_id)

    # s3 (originally last) should now be first
    expect_equal(ids_new, c(ids_orig[3], ids_orig[1], ids_orig[2]))
  })

  it("is a no-op when moving to same position", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    s1 <- prs$slides$add_slide(layout)
    s2 <- prs$slides$add_slide(layout)
    ids_orig <- sapply(prs$slides$to_list(), function(s) s$slide_id)

    prs$slides$move(s1, 1L)
    ids_new <- sapply(prs$slides$to_list(), function(s) s$slide_id)
    expect_equal(ids_new, ids_orig)
  })

  it("persists move across save/reload", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    s1 <- prs$slides$add_slide(layout)
    s2 <- prs$slides$add_slide(layout)
    s3 <- prs$slides$add_slide(layout)
    orig_id_last <- s3$slide_id

    prs$slides$move(s3, 1L)
    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    expect_equal(prs2$slides[[1]]$slide_id, orig_id_last)
  })

  it("errors on out-of-range index", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    s1 <- prs$slides$add_slide(layout)
    expect_error(prs$slides$move(s1, 5L), regexp = "out of range")
  })
})


# ============================================================================
# NotesSlide
# ============================================================================

describe("Slide$has_notes_slide", {
  it("returns FALSE when slide has no notes", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    expect_false(slide$has_notes_slide)
  })
})

describe("Slide$notes_slide", {
  it("creates a NotesSlide when accessed on a slide without notes", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    ns <- slide$notes_slide
    expect_s3_class(ns, "NotesSlide")
  })

  it("has_notes_slide is TRUE after accessing notes_slide", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    invisible(slide$notes_slide)
    expect_true(slide$has_notes_slide)
  })

  it("notes_text_frame returns TextFrame", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    ns  <- slide$notes_slide
    ntf <- ns$notes_text_frame
    expect_s3_class(ntf, "TextFrame")
  })

  it("notes_text_frame text is writable", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    ns  <- slide$notes_slide
    ntf <- ns$notes_text_frame
    ntf$text <- "My slide notes"
    expect_equal(ntf$text, "My slide notes")
  })

  it("round-trips through save/load", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    invisible(slide$notes_slide)
    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2  <- pptx_presentation(tmp)
    slide2 <- prs2$slides[[1]]
    expect_true(slide2$has_notes_slide)
  })
})


# ============================================================================
# Slide background
# ============================================================================

describe("Slide$background", {
  it("returns a .SlideBackground object", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    bg <- slide$background
    expect_s3_class(bg, ".SlideBackground")
  })

  it("background$fill returns FillFormat", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    ff <- slide$background$fill
    expect_s3_class(ff, "FillFormat")
  })

  it("can set a solid background colour", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    ff <- slide$background$fill
    ff$solid()
    ff$fore_color$rgb <- RGBColor(0xFF, 0x00, 0x00)
    expect_equal(as.character(slide$background$fill$fore_color$rgb), "FF0000")
  })

  it("solid background round-trips through save/load", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    slide$background$fill$solid()
    slide$background$fill$fore_color$rgb <- RGBColor(0x00, 0xFF, 0x00)
    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2  <- pptx_presentation(tmp)
    rgb2  <- prs2$slides[[1]]$background$fill$fore_color$rgb
    expect_equal(as.character(rgb2), "00FF00")
  })
})


# ============================================================================
# Slides$duplicate_slide
# ============================================================================

describe("Slides$duplicate_slide", {
  it("increases slide count by one", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    n_before <- length(prs$slides)
    prs$slides$duplicate_slide(slide)
    expect_equal(length(prs$slides), n_before + 1L)
  })

  it("returns a Slide object", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    dup <- prs$slides$duplicate_slide(slide)
    expect_s3_class(dup, "Slide")
  })

  it("duplicate has the same slide layout", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    dup <- prs$slides$duplicate_slide(slide)
    expect_equal(dup$slide_layout$name, slide$slide_layout$name)
  })

  it("duplicate preserves text added to the source", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    txb    <- slide$shapes$add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    txb$text_frame$text <- "Hello from source"
    dup <- prs$slides$duplicate_slide(slide)
    texts <- sapply(dup$shapes$to_list(), function(s) {
      tryCatch(s$text_frame$text, error = function(e) "")
    })
    expect_true(any(grepl("Hello from source", texts)))
  })

  it("duplicate is appended after source in slide order", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    s1 <- prs$slides$add_slide(layout)
    s2 <- prs$slides$add_slide(layout)
    dup <- prs$slides$duplicate_slide(s1)
    # dup should be last (index 3)
    expect_equal(prs$slides[[3]]$slide_id, dup$slide_id)
  })

  it("round-trips through save/load", {
    prs    <- pptx_presentation()
    layout <- prs$slide_layouts[[1]]
    slide  <- prs$slides$add_slide(layout)
    txb    <- slide$shapes$add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    txb$text_frame$text <- "Duplicate test"
    prs$slides$duplicate_slide(slide)
    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    expect_equal(length(prs2$slides), 2L)
    texts <- sapply(prs2$slides[[2]]$shapes$to_list(), function(s) {
      tryCatch(s$text_frame$text, error = function(e) "")
    })
    expect_true(any(grepl("Duplicate test", texts)))
  })
})
