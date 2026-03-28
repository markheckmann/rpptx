# Tests for oxml-text.R and text-text.R
#
# Phase 5: TextFrame, Paragraph, Run, Font — XML elements and proxy objects.

pptx_path <- function(name) system.file("test_files", name, package = "rpptx")

# Helper: open test_slides.pptx and return the textbox shape (index 4, "TextBox 6")
textbox_shape <- function() {
  prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
  slide <- prs$slides[[1]]
  slide$shapes[[4]]
}


# ============================================================================
# CT_RegularTextRun — a:r
# ============================================================================

describe("CT_RegularTextRun", {
  it("text getter returns run text", {
    shape <- textbox_shape()
    r <- shape$element$txBody$p_lst[[1]]$r_lst[[1]]
    expect_equal(r$text, "Test text")
  })

  it("text setter modifies run text", {
    shape <- textbox_shape()
    r <- shape$element$txBody$p_lst[[1]]$r_lst[[1]]
    r$text <- "Modified"
    expect_equal(r$text, "Modified")
  })

  it("get_or_add_rPr creates rPr if absent then returns it", {
    shape <- textbox_shape()
    p  <- shape$element$txBody$p_lst[[1]]
    r  <- p$r_lst[[1]]
    rPr <- r$get_or_add_rPr()
    expect_s3_class(rPr, "CT_TextCharacterProperties")
  })
})


# ============================================================================
# CT_TextParagraph — a:p
# ============================================================================

describe("CT_TextParagraph", {
  it("text getter concatenates run text", {
    shape <- textbox_shape()
    p <- shape$element$txBody$p_lst[[1]]
    expect_equal(p$text, "Test text")
  })

  it("text setter replaces content", {
    shape <- textbox_shape()
    p <- shape$element$txBody$p_lst[[1]]
    p$text <- "New content"
    expect_equal(p$text, "New content")
  })

  it("content_children returns runs only (no pPr/endParaRPr)", {
    shape <- textbox_shape()
    p <- shape$element$txBody$p_lst[[1]]
    children <- p$content_children
    classes  <- vapply(children, function(e) class(e)[1], character(1))
    expect_true(all(classes %in% c("CT_RegularTextRun", "CT_TextLineBreak")))
  })

  it("add_r appends a run with text", {
    shape <- textbox_shape()
    p <- shape$element$txBody$p_lst[[1]]
    r <- p$add_r("extra")
    expect_s3_class(r, "CT_RegularTextRun")
    expect_equal(r$text, "extra")
  })

  it("add_br appends a line break", {
    shape <- textbox_shape()
    p <- shape$element$txBody$p_lst[[1]]
    p$add_br()
    children <- p$content_children
    types    <- vapply(children, function(e) class(e)[1], character(1))
    expect_true("CT_TextLineBreak" %in% types)
  })

  it("append_text splits on newline into runs and breaks", {
    shape <- textbox_shape()
    p <- shape$element$txBody$p_lst[[1]]
    p$text <- ""  # clear
    p$append_text("abc\ndef")
    expect_equal(p$text, "abc\vdef")  # \n in append_text becomes \v (line break)
  })

  it("r_lst active binding returns list of CT_RegularTextRun", {
    shape <- textbox_shape()
    p <- shape$element$txBody$p_lst[[1]]
    runs <- p$r_lst
    expect_type(runs, "list")
    expect_true(length(runs) >= 1L)
    expect_s3_class(runs[[1]], "CT_RegularTextRun")
  })
})


# ============================================================================
# CT_TextBody — p:txBody
# ============================================================================

describe("CT_TextBody", {
  it("p_lst returns list of CT_TextParagraph", {
    shape <- textbox_shape()
    ps <- shape$element$txBody$p_lst
    expect_type(ps, "list")
    expect_true(length(ps) >= 1L)
    expect_s3_class(ps[[1]], "CT_TextParagraph")
  })

  it("bodyPr returns CT_TextBodyProperties", {
    shape <- textbox_shape()
    bodyPr <- shape$element$txBody$bodyPr
    expect_s3_class(bodyPr, "CT_TextBodyProperties")
  })

  it("add_p appends a new CT_TextParagraph", {
    shape <- textbox_shape()
    txBody <- shape$element$txBody
    n_before <- length(txBody$p_lst)
    p <- txBody$add_p()
    expect_s3_class(p, "CT_TextParagraph")
    expect_equal(length(txBody$p_lst), n_before + 1L)
  })

  it("clear_content removes all paragraphs", {
    shape <- textbox_shape()
    txBody <- shape$element$txBody
    txBody$clear_content()
    expect_equal(length(txBody$p_lst), 0L)
  })

  it("is_empty TRUE when single empty paragraph", {
    shape <- textbox_shape()
    txBody <- shape$element$txBody
    txBody$clear_content()
    txBody$add_p()
    expect_true(txBody$is_empty)
  })

  it("is_empty FALSE when paragraph has content", {
    shape <- textbox_shape()
    expect_false(shape$element$txBody$is_empty)
  })
})


# ============================================================================
# CT_TextBodyProperties — a:bodyPr
# ============================================================================

describe("CT_TextBodyProperties", {
  it("lIns default is 91440 EMU", {
    shape <- textbox_shape()
    # Use a fresh shape with default bodyPr (no lIns attr set)
    prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    bodyPr <- slide$shapes[[2]]$element$txBody$bodyPr
    expect_equal(as.integer(bodyPr$lIns), 91440L)
  })

  it("anchor read/write", {
    shape  <- textbox_shape()
    bodyPr <- shape$element$txBody$bodyPr
    bodyPr$anchor <- "t"
    expect_equal(bodyPr$anchor, "t")
    bodyPr$anchor <- NULL
    expect_null(bodyPr$anchor)
  })
})


# ============================================================================
# CT_TextParagraphProperties — a:pPr
# ============================================================================

describe("CT_TextParagraphProperties", {
  it("lvl defaults to 0", {
    shape <- textbox_shape()
    pPr <- shape$element$txBody$p_lst[[1]]$get_or_add_pPr()
    expect_equal(pPr$lvl, 0L)
  })

  it("algn read/write", {
    shape <- textbox_shape()
    pPr   <- shape$element$txBody$p_lst[[1]]$get_or_add_pPr()
    pPr$algn <- "ctr"
    expect_equal(pPr$algn, "ctr")
    pPr$algn <- NULL
    expect_null(pPr$algn)
  })

  it("line_spacing NULL when no lnSpc", {
    shape <- textbox_shape()
    pPr   <- shape$element$txBody$p_lst[[1]]$get_or_add_pPr()
    pPr$`_remove_lnSpc`()
    expect_null(pPr$line_spacing)
  })

  it("line_spacing round-trips as float (lines)", {
    shape <- textbox_shape()
    pPr   <- shape$element$txBody$p_lst[[1]]$get_or_add_pPr()
    pPr$line_spacing <- 1.5
    expect_equal(pPr$line_spacing, 1.5)
  })

  it("line_spacing round-trips as Length (points)", {
    shape <- textbox_shape()
    pPr   <- shape$element$txBody$p_lst[[1]]$get_or_add_pPr()
    pPr$line_spacing <- Pt(14)
    expect_equal(as_pt(pPr$line_spacing), 14)
  })

  it("space_before / space_after round-trip", {
    shape <- textbox_shape()
    pPr   <- shape$element$txBody$p_lst[[1]]$get_or_add_pPr()
    pPr$space_before <- Pt(6)
    pPr$space_after  <- Pt(3)
    expect_equal(as_pt(pPr$space_before), 6)
    expect_equal(as_pt(pPr$space_after),  3)
  })
})


# ============================================================================
# Font
# ============================================================================

describe("Font", {
  it("bold read/write", {
    shape <- textbox_shape()
    r     <- shape$element$txBody$p_lst[[1]]$r_lst[[1]]
    font  <- Font$new(r$get_or_add_rPr())
    font$bold <- TRUE
    expect_true(font$bold)
    font$bold <- NULL
    expect_null(font$bold)
  })

  it("italic read/write", {
    shape <- textbox_shape()
    r     <- shape$element$txBody$p_lst[[1]]$r_lst[[1]]
    font  <- Font$new(r$get_or_add_rPr())
    font$italic <- TRUE
    expect_true(font$italic)
  })

  it("size round-trips through Pt()", {
    shape <- textbox_shape()
    r     <- shape$element$txBody$p_lst[[1]]$r_lst[[1]]
    font  <- Font$new(r$get_or_add_rPr())
    font$size <- Pt(18)
    expect_equal(as_pt(font$size), 18)
  })

  it("size NULL when no sz attr", {
    shape <- textbox_shape()
    r     <- shape$element$txBody$p_lst[[1]]$r_lst[[1]]
    font  <- Font$new(r$get_or_add_rPr())
    font$size <- NULL
    expect_null(font$size)
  })

  it("name read/write", {
    shape <- textbox_shape()
    r     <- shape$element$txBody$p_lst[[1]]$r_lst[[1]]
    font  <- Font$new(r$get_or_add_rPr())
    font$name <- "Arial"
    expect_equal(font$name, "Arial")
    font$name <- NULL
    expect_null(font$name)
  })

  it("underline TRUE = 'sng', FALSE = 'none', NULL = inherit", {
    shape <- textbox_shape()
    r     <- shape$element$txBody$p_lst[[1]]$r_lst[[1]]
    font  <- Font$new(r$get_or_add_rPr())
    font$underline <- TRUE
    expect_true(font$underline)
    font$underline <- FALSE
    expect_false(font$underline)
    font$underline <- NULL
    expect_null(font$underline)
  })
})


# ============================================================================
# Run
# ============================================================================

describe("Run", {
  it("text getter returns run text", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    run   <- tf$paragraphs[[1]]$runs[[1]]
    expect_equal(run$text, "Test text")
  })

  it("text setter modifies run text", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    run   <- tf$paragraphs[[1]]$runs[[1]]
    run$text <- "Changed"
    expect_equal(run$text, "Changed")
  })

  it("font returns Font object", {
    shape <- textbox_shape()
    run   <- shape$text_frame$paragraphs[[1]]$runs[[1]]
    expect_s3_class(run$font, "Font")
  })

  it("chained font assignment works (R6 write-back no-op)", {
    shape <- textbox_shape()
    run   <- shape$text_frame$paragraphs[[1]]$runs[[1]]
    run$font$bold <- TRUE
    expect_true(run$font$bold)
  })
})


# ============================================================================
# Paragraph
# ============================================================================

describe("Paragraph", {
  it("text getter returns paragraph text", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    expect_equal(para$text, "Test text")
  })

  it("text setter replaces content", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    para$text <- "Replaced"
    expect_equal(para$text, "Replaced")
  })

  it("runs returns list of Run objects", {
    shape <- textbox_shape()
    runs  <- shape$text_frame$paragraphs[[1]]$runs
    expect_type(runs, "list")
    expect_true(length(runs) >= 1L)
    expect_s3_class(runs[[1]], "Run")
  })

  it("font returns Font object (from a:defRPr)", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    expect_s3_class(para$font, "Font")
  })

  it("add_run appends a Run", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    run   <- para$add_run()
    run$text <- "new run"
    expect_equal(para$text, "Test textnew run")
  })

  it("add_line_break appends a line break", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    para$add_line_break()
    expect_true(grepl("\v", para$text))
  })

  it("clear removes all content", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    para$clear()
    expect_equal(para$text, "")
    expect_length(para$runs, 0L)
  })

  it("alignment read/write", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    para$alignment <- "ctr"
    expect_equal(para$alignment, "ctr")
  })

  it("level default is 0", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    expect_equal(para$level, 0L)
  })

  it("line_spacing read/write (float = lines)", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    para$line_spacing <- 2.0
    expect_equal(para$line_spacing, 2.0)
  })

  it("line_spacing read/write (Length = fixed)", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    para$line_spacing <- Pt(12)
    expect_equal(as_pt(para$line_spacing), 12)
  })

  it("space_before / space_after read/write", {
    shape <- textbox_shape()
    para  <- shape$text_frame$paragraphs[[1]]
    para$space_before <- Pt(6)
    para$space_after  <- Pt(3)
    expect_equal(as_pt(para$space_before), 6)
    expect_equal(as_pt(para$space_after),  3)
  })
})


# ============================================================================
# TextFrame
# ============================================================================

describe("TextFrame", {
  it("paragraphs returns list of Paragraph objects", {
    shape <- textbox_shape()
    paras <- shape$text_frame$paragraphs
    expect_type(paras, "list")
    expect_true(length(paras) >= 1L)
    expect_s3_class(paras[[1]], "Paragraph")
  })

  it("text getter joins paragraphs with newline", {
    shape <- textbox_shape()
    expect_equal(shape$text_frame$text, "Test text")
  })

  it("text setter splits on newline into paragraphs", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    tf$text <- "line1\nline2\nline3"
    expect_equal(tf$text, "line1\nline2\nline3")
    expect_equal(length(tf), 3L)
  })

  it("add_paragraph appends a new Paragraph", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    n_before <- length(tf)
    p <- tf$add_paragraph()
    expect_s3_class(p, "Paragraph")
    expect_equal(length(tf), n_before + 1L)
  })

  it("clear removes all but first paragraph and clears it", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    tf$text <- "a\nb\nc"
    tf$clear()
    expect_equal(length(tf), 1L)
    expect_equal(tf$text, "")
  })

  it("[[ returns Paragraph at 1-based index", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    tf$text <- "first\nsecond"
    expect_equal(tf[[1]]$text, "first")
    expect_equal(tf[[2]]$text, "second")
  })

  it("length() returns paragraph count", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    tf$text <- "a\nb"
    expect_equal(length(tf), 2L)
  })

  it("word_wrap read/write", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    tf$word_wrap <- TRUE
    expect_true(tf$word_wrap)
    tf$word_wrap <- FALSE
    expect_false(tf$word_wrap)
    tf$word_wrap <- NULL
    expect_null(tf$word_wrap)
  })

  it("vertical_anchor read/write", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    tf$vertical_anchor <- "t"
    expect_equal(tf$vertical_anchor, "t")
  })

  it("margin_left/right/top/bottom read/write", {
    shape <- textbox_shape()
    tf    <- shape$text_frame
    tf$margin_left   <- Inches(0.1)
    tf$margin_right  <- Inches(0.1)
    tf$margin_top    <- Inches(0.05)
    tf$margin_bottom <- Inches(0.05)
    expect_equal(as_inches(tf$margin_left),   0.1)
    expect_equal(as_inches(tf$margin_right),  0.1)
    expect_equal(as_inches(tf$margin_top),    0.05)
    expect_equal(as_inches(tf$margin_bottom), 0.05)
  })
})


# ============================================================================
# Shape$text_frame and Shape$text
# ============================================================================

describe("Shape text_frame and text", {
  it("text_frame returns TextFrame", {
    shape <- textbox_shape()
    expect_s3_class(shape$text_frame, "TextFrame")
  })

  it("text getter returns shape text", {
    shape <- textbox_shape()
    expect_equal(shape$text, "Test text")
  })

  it("text setter modifies shape text", {
    shape <- textbox_shape()
    shape$text <- "Updated"
    expect_equal(shape$text, "Updated")
  })

  it("chained shape$text_frame$text setter works", {
    shape <- textbox_shape()
    shape$text_frame$text <- "Chained"
    expect_equal(shape$text, "Chained")
  })

  it("get_or_add_txBody creates txBody on placeholder without one", {
    prs   <- pptx_presentation(pptx_path("test_slides.pptx"))
    slide <- prs$slides[[1]]
    title_shape <- slide$shapes[[1]]
    expect_s3_class(title_shape$text_frame, "TextFrame")
  })
})
