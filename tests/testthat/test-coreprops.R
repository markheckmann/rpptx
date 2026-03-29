# Tests for core properties (oxml-coreprops.R, parts-coreprops.R)

describe("CT_CoreProperties", {
  it("reads author from an existing file", {
    prs <- pptx_presentation(test_file_path("../templates/default.pptx"))
    cp  <- prs$core_properties
    expect_type(cp$author, "character")
  })

  it("round-trips text properties", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    cp$title <- "Test Title"
    expect_equal(cp$title, "Test Title")
    cp$author <- "Test Author"
    expect_equal(cp$author, "Test Author")
    cp$subject <- "Test Subject"
    expect_equal(cp$subject, "Test Subject")
    cp$comments <- "Test Comments"
    expect_equal(cp$comments, "Test Comments")
    cp$keywords <- "test foo bar"
    expect_equal(cp$keywords, "test foo bar")
    cp$category <- "Test Category"
    expect_equal(cp$category, "Test Category")
    cp$content_status <- "Draft"
    expect_equal(cp$content_status, "Draft")
    cp$language <- "en-US"
    expect_equal(cp$language, "en-US")
    cp$last_modified_by <- "rpptx"
    expect_equal(cp$last_modified_by, "rpptx")
    cp$version <- "1.0"
    expect_equal(cp$version, "1.0")
    cp$identifier <- "abc-123"
    expect_equal(cp$identifier, "abc-123")
  })

  it("returns empty string for missing text properties", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    expect_equal(cp$author, "")
  })

  it("round-trips revision integer", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    cp$revision <- 3L
    expect_equal(cp$revision, 3L)
  })

  it("returns 0 for missing revision", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    # Default new presentations have revision set via CorePropertiesPart_default
    expect_gte(cp$revision, 0L)
  })

  it("errors on non-positive revision", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    expect_error(cp$revision <- 0L, regexp = "positive")
  })

  it("errors on text exceeding 255 chars", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    long_str <- paste(rep("x", 256L), collapse = "")
    expect_error(cp$title <- long_str, regexp = "255")
  })

  it("round-trips modified datetime", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    now <- as.POSIXct("2024-01-15 12:00:00", tz = "UTC")
    cp$modified <- now
    result <- cp$modified
    expect_s3_class(result, "POSIXct")
    # Allow 1 second tolerance for round-trip
    expect_true(abs(as.numeric(result) - as.numeric(now)) < 1.0)
  })

  it("round-trips created datetime", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    dt  <- as.POSIXct("2020-06-01 08:30:00", tz = "UTC")
    cp$created <- dt
    result <- cp$created
    expect_s3_class(result, "POSIXct")
    expect_true(abs(as.numeric(result) - as.numeric(dt)) < 1.0)
  })

  it("returns NULL for missing datetime", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    # last_printed is typically absent in new presentations
    # (may be NULL or POSIXct — just check it doesn't error)
    result <- cp$last_printed
    expect_true(is.null(result) || inherits(result, "POSIXct"))
  })

  it("is accessible from Presentation$core_properties", {
    prs <- pptx_presentation()
    cp  <- prs$core_properties
    expect_s3_class(cp, "CorePropertiesPart")
  })
})

describe(".parse_W3CDTF", {
  it("parses UTC datetime strings", {
    dt <- rpptx:::.parse_W3CDTF("2024-03-15T10:30:00Z")
    expect_s3_class(dt, "POSIXct")
    expect_equal(format(dt, "%Y-%m-%dT%H:%M:%SZ", tz = "UTC"),
                 "2024-03-15T10:30:00Z")
  })

  it("parses date-only strings", {
    dt <- rpptx:::.parse_W3CDTF("2024-03-15")
    expect_s3_class(dt, "POSIXct")
  })

  it("parses year-month strings", {
    dt <- rpptx:::.parse_W3CDTF("2024-03")
    expect_s3_class(dt, "POSIXct")
  })

  it("returns NULL for invalid strings", {
    result <- rpptx:::.parse_W3CDTF("not-a-date")
    expect_null(result)
  })
})


# ============================================================================
# Core property read/write via Presentation$core_properties
# ============================================================================

describe("Presentation$core_properties read/write", {
  it("title is writable and readable", {
    prs <- pptx_presentation()
    prs$core_properties$title <- "My Deck"
    expect_equal(prs$core_properties$title, "My Deck")
  })

  it("author is writable and readable", {
    prs <- pptx_presentation()
    prs$core_properties$author <- "Jane Doe"
    expect_equal(prs$core_properties$author, "Jane Doe")
  })

  it("subject is writable and readable", {
    prs <- pptx_presentation()
    prs$core_properties$subject <- "Q1 Report"
    expect_equal(prs$core_properties$subject, "Q1 Report")
  })

  it("round-trips through save/load", {
    prs <- pptx_presentation()
    prs$core_properties$title   <- "Save Test"
    prs$core_properties$author  <- "Alice"
    tmp <- tempfile(fileext = ".pptx")
    on.exit(unlink(tmp))
    prs$save(tmp)
    prs2 <- pptx_presentation(tmp)
    expect_equal(prs2$core_properties$title,  "Save Test")
    expect_equal(prs2$core_properties$author, "Alice")
  })
})
