# Tests for opc-package.R — OpcPackage, Part, Relationships, etc.

describe("OpcPackage", {
  default_pptx <- test_file_path("../templates/default.pptx")

  it("opens a .pptx file and enumerates parts", {
    pkg <- OpcPackage_open(default_pptx)
    parts <- pkg$iter_parts()
    expect_true(length(parts) > 0)
    # Default template should have presentation, theme, slide layouts, etc.
    partnames <- vapply(parts, function(p) as.character(p$partname), character(1))
    expect_true("/ppt/presentation.xml" %in% partnames)
    expect_true("/ppt/theme/theme1.xml" %in% partnames)
  })

  it("provides main_document_part", {
    pkg <- OpcPackage_open(default_pptx)
    main_part <- pkg$main_document_part
    expect_true(inherits(main_part, "Part"))
    expect_equal(as.character(main_part$partname), "/ppt/presentation.xml")
  })

  it("enumerates all relationships depth-first", {
    pkg <- OpcPackage_open(default_pptx)
    rels <- pkg$iter_rels()
    expect_true(length(rels) > 0)
    # Every rel should have an rId
    for (rel in rels) {
      expect_true(grepl("^rId\\d+$", rel$rId))
    }
  })

  it("round-trips a .pptx file", {
    pkg <- OpcPackage_open(default_pptx)
    out_file <- tempfile(fileext = ".pptx")
    on.exit(unlink(out_file), add = TRUE)

    pkg$save(out_file)
    expect_true(file.exists(out_file))
    expect_true(file.info(out_file)$size > 0)

    # Reopen and compare partnames
    pkg2 <- OpcPackage_open(out_file)
    pn1 <- sort(vapply(pkg$iter_parts(), function(p) as.character(p$partname), character(1)))
    pn2 <- sort(vapply(pkg2$iter_parts(), function(p) as.character(p$partname), character(1)))
    expect_equal(pn1, pn2)
  })

  it("computes next_partname correctly", {
    pkg <- OpcPackage_open(default_pptx)
    # No slides in default template → slide1 should be next
    next_pn <- pkg$next_partname("/ppt/slides/slide%d.xml")
    expect_s3_class(next_pn, "PackURI")
    expect_equal(as.character(next_pn), "/ppt/slides/slide1.xml")
  })
})


describe("Part", {
  it("stores partname, content_type, and blob", {
    pn <- PackURI("/ppt/slides/slide1.xml")
    part <- Part$new(pn, "application/xml", NULL, charToRaw("<slide/>"))
    expect_equal(as.character(part$partname), "/ppt/slides/slide1.xml")
    expect_equal(part$content_type, "application/xml")
    expect_equal(rawToChar(part$blob), "<slide/>")
  })

  it("validates partname type on set", {
    pn <- PackURI("/ppt/slides/slide1.xml")
    part <- Part$new(pn, "application/xml", NULL)
    expect_error(part$partname <- "not a PackURI", "must be a PackURI")
  })

  it("returns empty raw for missing blob", {
    pn <- PackURI("/test.xml")
    part <- Part$new(pn, "application/xml", NULL)
    expect_equal(part$blob, raw(0))
  })
})


describe("Relationships", {
  it("starts empty", {
    rels <- Relationships$new("/ppt")
    expect_equal(length(rels), 0)
    expect_equal(length(rels$values()), 0)
  })

  it("can add and retrieve relationships", {
    rels <- Relationships$new("/")
    pn <- PackURI("/ppt/presentation.xml")
    target <- Part$new(pn, "application/xml", NULL)

    rId <- rels$get_or_add(RT$OFFICE_DOCUMENT, target)
    expect_equal(rId, "rId1")
    expect_equal(length(rels), 1)

    rel <- rels$get(rId)
    expect_false(rel$is_external)
    expect_equal(rel$reltype, RT$OFFICE_DOCUMENT)
  })

  it("returns existing rId for duplicate relationship", {
    rels <- Relationships$new("/")
    target <- Part$new(PackURI("/ppt/presentation.xml"), "application/xml", NULL)

    rId1 <- rels$get_or_add(RT$OFFICE_DOCUMENT, target)
    rId2 <- rels$get_or_add(RT$OFFICE_DOCUMENT, target)
    expect_equal(rId1, rId2)
    expect_equal(length(rels), 1)
  })

  it("finds part by reltype", {
    rels <- Relationships$new("/")
    target <- Part$new(PackURI("/ppt/presentation.xml"), "application/xml", NULL)
    rels$get_or_add(RT$OFFICE_DOCUMENT, target)

    found <- rels$part_with_reltype(RT$OFFICE_DOCUMENT)
    expect_identical(found, target)
  })

  it("errors on missing reltype", {
    rels <- Relationships$new("/")
    expect_error(rels$part_with_reltype(RT$SLIDE), "no relationship of type")
  })

  it("can add external relationships", {
    rels <- Relationships$new("/ppt/slides")
    rId <- rels$get_or_add_ext_rel(RT$HYPERLINK, "https://example.com")
    expect_equal(rId, "rId1")

    rel <- rels$get(rId)
    expect_true(rel$is_external)
    expect_equal(rel$target_ref, "https://example.com")
  })

  it("generates xml_bytes", {
    rels <- Relationships$new("/")
    target <- Part$new(PackURI("/ppt/presentation.xml"), "application/xml", NULL)
    rels$get_or_add(RT$OFFICE_DOCUMENT, target)

    bytes <- rels$xml_bytes()
    xml_str <- rawToChar(bytes)
    expect_true(grepl("Relationships", xml_str))
    expect_true(grepl("Relationship", xml_str))
    expect_true(grepl("rId1", xml_str))
  })

  it("can pop a relationship", {
    rels <- Relationships$new("/")
    target <- Part$new(PackURI("/ppt/presentation.xml"), "application/xml", NULL)
    rId <- rels$get_or_add(RT$OFFICE_DOCUMENT, target)

    popped <- rels$pop(rId)
    expect_equal(popped$rId, rId)
    expect_equal(length(rels), 0)
  })
})


describe("Relationship", {
  it("exposes properties for internal relationship", {
    target <- Part$new(PackURI("/ppt/presentation.xml"), "application/xml", NULL)
    rel <- Relationship$new("/", "rId1", RT$OFFICE_DOCUMENT, RTM$INTERNAL, target)

    expect_equal(rel$rId, "rId1")
    expect_equal(rel$reltype, RT$OFFICE_DOCUMENT)
    expect_false(rel$is_external)
    expect_identical(rel$target_part, target)
    expect_equal(rel$target_ref, "ppt/presentation.xml")
  })

  it("exposes properties for external relationship", {
    rel <- Relationship$new("/ppt/slides", "rId1", RT$HYPERLINK, RTM$EXTERNAL, "https://example.com")

    expect_true(rel$is_external)
    expect_equal(rel$target_ref, "https://example.com")
    expect_error(rel$target_part, "undefined for external")
    expect_error(rel$target_partname, "undefined for external")
  })
})


describe("ContentTypeMap", {
  it("resolves content types from XML", {
    pkg <- OpcPackage_open(test_file_path("../templates/default.pptx"))
    pr <- PackageReader$new(test_file_path("../templates/default.pptx"))
    ct_blob <- pr$get_blob(CONTENT_TYPES_URI)
    ct_map <- ContentTypeMap$from_xml(ct_blob)

    # Override partname lookup
    ct <- ct_map$get(PackURI("/ppt/presentation.xml"))
    expect_true(grepl("presentation", ct))

    # Default extension lookup
    ct_xml <- ct_map$get(PackURI("/some/file.xml"))
    expect_equal(ct_xml, "application/xml")
  })

  it("errors on unknown partname", {
    pr <- PackageReader$new(test_file_path("../templates/default.pptx"))
    ct_blob <- pr$get_blob(CONTENT_TYPES_URI)
    ct_map <- ContentTypeMap$from_xml(ct_blob)

    expect_error(ct_map$get(PackURI("/unknown.xyz")), "no content-type")
  })
})
