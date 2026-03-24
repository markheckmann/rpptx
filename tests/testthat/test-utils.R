# Tests for R/utils.R -- Length classes and unit conversions.

describe("Length", {
  it("creates a Length from EMU value", {
    l <- Length(914400L)
    expect_s3_class(l, "Length")
    expect_equal(as.integer(l), 914400L)
  })

  it("prints nicely", {
    l <- Inches(1)
    expect_output(print(l), "Length:.*914400 EMU")
  })
})

describe("Inches", {
  it("converts inches to EMU", {
    expect_equal(as.integer(Inches(1)), 914400L)
    expect_equal(as.integer(Inches(2)), 1828800L)
  })

  it("handles fractional inches", {
    expect_equal(as.integer(Inches(0.5)), 457200L)
  })
})

describe("Cm", {
  it("converts centimeters to EMU", {
    expect_equal(as.integer(Cm(1)), 360000L)
    expect_equal(as.integer(Cm(2.54)), as.integer(Inches(1)))
  })
})

describe("Mm", {
  it("converts millimeters to EMU", {
    expect_equal(as.integer(Mm(1)), 36000L)
    expect_equal(as.integer(Mm(10)), as.integer(Cm(1)))
  })
})

describe("Pt", {
  it("converts points to EMU", {
    expect_equal(as.integer(Pt(1)), 12700L)
    expect_equal(as.integer(Pt(72)), as.integer(Inches(1)))
  })
})

describe("Emu", {
  it("passes EMU value through", {
    expect_equal(as.integer(Emu(914400)), 914400L)
  })
})

describe("Centipoints", {
  it("converts centipoints to EMU", {
    expect_equal(as.integer(Centipoints(1)), 127L)
    expect_equal(as.integer(Centipoints(100)), as.integer(Pt(1)))
  })
})

describe("as_inches", {
  it("converts EMU to inches", {
    expect_equal(as_inches(Inches(1)), 1.0)
    expect_equal(as_inches(Inches(2.5)), 2.5)
  })
})

describe("as_cm", {
  it("converts EMU to centimeters", {
    expect_equal(as_cm(Cm(5)), 5.0)
  })
})

describe("as_pt", {
  it("converts EMU to points", {
    expect_equal(as_pt(Pt(12)), 12.0)
  })
})

describe("as_centipoints", {
  it("converts EMU to centipoints (integer division)", {
    expect_equal(as_centipoints(Pt(1)), 100L)
    expect_equal(as_centipoints(Centipoints(150)), 150L)
  })
})

describe("lazy_active_binding", {
  it("caches the result of the function", {
    call_count <- 0L
    TestClass <- R6::R6Class("TestClass",
      private = list(
        .cached_value = NULL
      ),
      active = list(
        value = lazy_active_binding(
          fn = function(self) {
            call_count <<- call_count + 1L
            42
          },
          cache_field = ".cached_value"
        )
      )
    )
    obj <- TestClass$new()
    expect_equal(obj$value, 42)
    expect_equal(obj$value, 42)
    expect_equal(call_count, 1L)  # only called once
  })

  it("raises error on assignment", {
    TestClass <- R6::R6Class("TestClass",
      private = list(.cached = NULL),
      active = list(
        value = lazy_active_binding(
          fn = function(self) 42,
          cache_field = ".cached"
        )
      )
    )
    obj <- TestClass$new()
    expect_error(obj$value <- 99, "read-only")
  })
})
