# Tests for the Presentation API (Phase 2)

describe("pptx_presentation", {
  it("creates a presentation from the default template", {
    prs <- pptx_presentation()
    expect_s3_class(prs, "Presentation")
  })

  it("opens a .pptx file", {
    prs <- pptx_presentation(test_file_path("../templates/default.pptx"))
    expect_s3_class(prs, "Presentation")
  })

  it("opens a minimal test file", {
    prs <- pptx_presentation(test_file_path("minimal.pptx"))
    expect_s3_class(prs, "Presentation")
  })

  it("errors on invalid file", {
    fake <- tempfile(fileext = ".pptx")
    writeLines("not a zip", fake)
    on.exit(unlink(fake), add = TRUE)
    expect_error(suppressWarnings(pptx_presentation(fake)))
  })
})


describe("Presentation$slide_width/slide_height", {
  it("returns slide dimensions in EMU", {
    prs <- pptx_presentation()
    # Default template: 10" x 7.5" (standard 4:3)
    expect_equal(prs$slide_width, Emu(9144000))
    expect_equal(prs$slide_height, Emu(6858000))
    expect_equal(as_inches(prs$slide_width), 10)
    expect_equal(as_inches(prs$slide_height), 7.5)
  })

  it("can set slide dimensions", {
    prs <- pptx_presentation()
    prs$slide_width <- Inches(13.333)
    prs$slide_height <- Inches(7.5)
    expect_equal(prs$slide_width, Inches(13.333))
  })
})


describe("Presentation$save", {
  it("round-trips a presentation", {
    prs <- pptx_presentation()
    out <- tempfile(fileext = ".pptx")
    on.exit(unlink(out), add = TRUE)

    prs$save(out)
    expect_true(file.exists(out))

    prs2 <- pptx_presentation(out)
    expect_equal(prs2$slide_width, prs$slide_width)
    expect_equal(prs2$slide_height, prs$slide_height)
  })

  it("preserves modified dimensions", {
    prs <- pptx_presentation()
    prs$slide_width <- Inches(16)
    prs$slide_height <- Inches(9)

    out <- tempfile(fileext = ".pptx")
    on.exit(unlink(out), add = TRUE)

    prs$save(out)
    prs2 <- pptx_presentation(out)
    expect_equal(prs2$slide_width, Inches(16))
    expect_equal(prs2$slide_height, Inches(9))
  })
})


describe("Presentation$element", {
  it("returns a CT_Presentation element", {
    prs <- pptx_presentation()
    expect_s3_class(prs$element, "CT_Presentation")
  })
})


describe("Presentation$part", {
  it("returns a PresentationPart", {
    prs <- pptx_presentation()
    expect_s3_class(prs$part, "PresentationPart")
  })
})
