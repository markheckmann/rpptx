# Tests for opc-serialized.R — PackageReader, PackageWriter

describe("PackageReader", {
  default_pptx <- test_file_path("../templates/default.pptx")

  it("reads blobs from a .pptx file", {
    pr <- PackageReader$new(default_pptx)
    expect_true(pr$contains(CONTENT_TYPES_URI))
    expect_true(pr$contains(PackURI("/ppt/presentation.xml")))
  })

  it("returns raw bytes for a blob", {
    pr <- PackageReader$new(default_pptx)
    blob <- pr$get_blob(CONTENT_TYPES_URI)
    expect_type(blob, "raw")
    expect_true(length(blob) > 0)
    # Should be XML
    xml_str <- rawToChar(blob)
    expect_true(grepl("<Types", xml_str))
  })

  it("returns rels XML for parts that have them", {
    pr <- PackageReader$new(default_pptx)
    # Package root has rels
    rels_xml <- pr$rels_xml_for(PACKAGE_URI)
    expect_type(rels_xml, "raw")
    expect_true(grepl("Relationships", rawToChar(rels_xml)))
  })

  it("returns NULL for parts without rels", {
    pr <- PackageReader$new(default_pptx)
    # theme usually has no rels of its own
    fake_uri <- PackURI("/nonexistent.xml")
    expect_null(pr$rels_xml_for(fake_uri))
  })

  it("errors on non-existent part", {
    pr <- PackageReader$new(default_pptx)
    expect_error(pr$get_blob(PackURI("/does_not_exist.xml")), "no member")
  })
})


describe("PackageWriter", {
  it("writes a valid ZIP package", {
    out_file <- tempfile(fileext = ".pptx")
    on.exit(unlink(out_file), add = TRUE)

    # Create minimal content
    rels <- Relationships$new(pack_uri_base(PACKAGE_URI))
    pn <- PackURI("/ppt/presentation.xml")
    part <- Part$new(pn, CT$PML_PRESENTATION_MAIN, NULL,
                     charToRaw("<presentation/>"))

    rels$get_or_add(RT$OFFICE_DOCUMENT, part)

    writer <- PackageWriter$new()
    writer$write(out_file, rels, list(part))

    expect_true(file.exists(out_file))
    expect_true(file.info(out_file)$size > 0)

    # Verify we can read it back
    pr <- PackageReader$new(out_file)
    expect_true(pr$contains(CONTENT_TYPES_URI))
    expect_true(pr$contains(pn))
  })
})
