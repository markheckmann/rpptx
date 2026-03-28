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
