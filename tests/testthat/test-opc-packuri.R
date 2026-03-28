# Tests for opc-packuri.R — PackURI and related functions

describe("PackURI", {
  it("creates a PackURI from a valid string", {
    uri <- PackURI("/ppt/slides/slide1.xml")
    expect_s3_class(uri, "PackURI")
    expect_equal(as.character(uri), "/ppt/slides/slide1.xml")
  })

  it("errors if string does not start with /", {
    expect_error(PackURI("ppt/slides/slide1.xml"), "must begin with slash")
  })

  it("errors if input is not a single string", {
    expect_error(PackURI(42), "must be a single character string")
    expect_error(PackURI(c("/a", "/b")), "must be a single character string")
  })
})

describe("pack_uri_base", {
  it("returns the directory portion", {
    expect_equal(pack_uri_base(PackURI("/ppt/slides/slide1.xml")), "/ppt/slides")
  })

  it("returns / for root-level files", {
    expect_equal(pack_uri_base(PackURI("/[Content_Types].xml")), "/")
  })
})

describe("pack_uri_ext", {
  it("returns the file extension", {
    expect_equal(pack_uri_ext(PackURI("/ppt/slides/slide1.xml")), "xml")
    expect_equal(pack_uri_ext(PackURI("/docProps/thumbnail.jpeg")), "jpeg")
  })
})

describe("pack_uri_filename", {
  it("returns the filename", {
    expect_equal(pack_uri_filename(PackURI("/ppt/slides/slide1.xml")), "slide1.xml")
  })
})

describe("pack_uri_idx", {
  it("returns the trailing integer", {
    expect_equal(pack_uri_idx(PackURI("/ppt/slides/slide21.xml")), 21L)
    expect_equal(pack_uri_idx(PackURI("/ppt/slides/slide1.xml")), 1L)
  })

  it("returns NULL for non-array partnames", {
    expect_null(pack_uri_idx(PackURI("/ppt/presentation.xml")))
  })
})

describe("pack_uri_membername", {
  it("strips leading slash", {
    expect_equal(pack_uri_membername(PackURI("/ppt/slides/slide1.xml")),
                 "ppt/slides/slide1.xml")
  })
})

describe("pack_uri_relative_ref", {
  it("computes relative reference from base", {
    uri <- PackURI("/ppt/slideLayouts/slideLayout1.xml")
    expect_equal(pack_uri_relative_ref(uri, "/ppt/slides"),
                 "../slideLayouts/slideLayout1.xml")
  })

  it("handles root base URI", {
    uri <- PackURI("/ppt/presentation.xml")
    expect_equal(pack_uri_relative_ref(uri, "/"), "ppt/presentation.xml")
  })
})

describe("pack_uri_rels_uri", {
  it("returns the .rels URI for a part", {
    uri <- PackURI("/ppt/slides/slide1.xml")
    rels <- pack_uri_rels_uri(uri)
    expect_equal(as.character(rels), "/ppt/slides/_rels/slide1.xml.rels")
  })

  it("handles the package root URI", {
    rels <- pack_uri_rels_uri(PACKAGE_URI)
    expect_equal(as.character(rels), "/_rels/.rels")
  })
})

describe("pack_uri_from_rel_ref", {
  it("resolves a relative reference to a PackURI", {
    result <- pack_uri_from_rel_ref("/ppt/slides", "../slideLayouts/slideLayout1.xml")
    expect_equal(as.character(result), "/ppt/slideLayouts/slideLayout1.xml")
  })

  it("resolves from root base", {
    result <- pack_uri_from_rel_ref("/", "ppt/presentation.xml")
    expect_equal(as.character(result), "/ppt/presentation.xml")
  })
})
