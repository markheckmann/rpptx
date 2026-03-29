# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working
with code in this repository.

## Development Commands

``` r
devtools::load_all(".")       # Load package for interactive use
devtools::document()          # Regenerate NAMESPACE and man/ from roxygen2
devtools::test()              # Run full test suite
devtools::check()             # R CMD check (equivalent of CI)

# Run a single test file
testthat::test_file("tests/testthat/test-utils.R")

# Run tests matching a pattern
devtools::test(filter = "text")
```

The CI runs
`rcmdcheck::rcmdcheck(args = c("--no-manual", "--as-cran"), error_on = "warning")`.
Any WARNING is a build failure.

## Architecture

The package is a faithful R port of python-pptx. It mirrors the same
4-layer architecture:

    pptx_presentation()           ← public entry point (api.R)
        │
        ▼
    Domain Layer                  ← presentation.R, slide.R, shapes-*.R, text-text.R,
                                     table.R, chart-*.R, dml-*.R, action.R
        │ wraps / navigates
        ▼
    Parts Layer                   ← parts-*.R  (one file per part type)
        │ owns XML trees
        ▼
    OXML Layer                    ← oxml-*.R   (CT_* element classes)
        │ I/O via xml2
        ▼
    OPC Layer                     ← opc-*.R    (ZIP, relationships, content types)

**OPC layer** (`opc-*.R`): Pure ZIP/packaging — no business logic.
`PackURI` for part paths, `OpcPackage` for reading/writing,
`Relationships` for part-to-part edges.

**OXML layer** (`oxml-*.R`): Direct 1-to-1 mapping to OOXML schema. Each
XML element type has a `CT_*` class. No business logic here either.
Built on `BaseOxmlElement` which wraps an xml2 node.

**Parts layer** (`parts-*.R`): OPC parts that own an XML tree and
relationships. Bridge between packaging and content. E.g., `SlidePart`
holds a `CT_Slide` element and knows its layout part relationship.

**Domain layer**: User-facing R6 classes wrapping OXML elements via
`$._element`. Active bindings expose properties. Collections (`Slides`,
`SlideShapes`, etc.) support `[[`,
[`length()`](https://rdrr.io/r/base/length.html) via S3 methods.

## The xmlchemy System

The core of the OXML layer is in `oxml-xmlchemy.R` and `oxml-init.R`.

**`BaseOxmlElement`** wraps an xml2 node with helpers:
[`find()`](https://rdrr.io/r/utils/apropos.html), `findall()`,
`get_attr()`, `set_attr()`, `insert_element_before()`, `xpath()`. The
`$tag` active binding returns the Clark-notation tag `{uri}localname`.

**`define_oxml_element()`** is a factory that generates R6 classes for
XML elements with child and attribute descriptors:

``` r
CT_Example <- define_oxml_element(
  classname  = "CT_Example",
  tag        = "p:example",
  children   = list(
    title = zero_or_one("p:title", successors = c("p:body")),
    items = zero_or_more("p:item")
  ),
  attributes = list(
    val = optional_attribute("val", XsdString, default = NULL)
  )
)
```

This generates active bindings (`$title`, `$items`, `$val`) and
`get_or_add_*()` / `_remove_*()` methods automatically.

**Element registry**: `register_element_cls("p:example", CT_Example)`
registers a class for a tag. `wrap_element(xml2_node)` looks up the tag
and returns the correct R6 instance. Called throughout the codebase when
reading children from the XML tree.

**Child specs**: - `zero_or_one()` — optional, returns element or NULL;
auto-creates on `get_or_add_*()` - `zero_or_more()` / `one_or_more()` —
returns list; modified in-place - `one_and_only_one()` — required,
errors if missing

**Attribute specs**: `optional_attribute(name, simple_type, default)`,
`required_attribute(name, simple_type)`. Simple types (`XsdString`,
`XsdInt`, `XsdBool`, `XsdDouble`) are in `oxml-simpletypes.R`.

## Key Patterns & Pitfalls

### Active binding write-back (replacement chain)

`obj$a$b <- value` triggers R’s replacement chain: it calls `obj$a`
(read), then tries `obj$a <- modified_a` (write-back). Active bindings
that return sub-objects need a no-op setter to absorb this:

``` r
active = list(
  font = function(value) {
    if (!missing(value)) return(invisible(NULL))   # no-op write-back
    Font$new(private$.element$get_or_add_rPr())
  }
)
```

Collections (`[[<-` S3 methods) need the same treatment. Missing no-op
setters cause silent failures or errors.

### Never use `func()$field <- value`

`m$.add_size()$val <- value` is a bad assignment — R cannot write back
to a function return value. Always split into two statements:

``` r
sz <- m$.add_size()
sz$val <- value
```

### OOXML schema ordering

When inserting elements into axis or other complex elements, order
matters. Use `insert_element_before()` with the correct successor tag.
Example: `<c:txPr>` must precede `<c:crossAx>` inside axis elements.

### @noRd vs @export

- Internal R6 classes (CT\_*, Base*, Parts\*): `@noRd` only, no
  `@export` — they appear in the NAMESPACE only if users need to
  instantiate them directly.
- User-facing classes users instantiate directly (`CategoryChartData`,
  `XyChartData`, `BubbleChartData`, `FreeformBuilder`): `@export` with a
  title+description block (no `@noRd`).
- Functions and enum constants: `@export` with full docs.
- Never use `@export @noRd` together — exports without docs fail R CMD
  check with “Undocumented code objects” WARNING.

### Namespace map

`opc-constants.R` defines `.nsmap` (prefix → URI) and `.pfxmap` (URI →
prefix). Always use `qn("p:sp")` to get Clark-notation names, or pass
`ns = c(p = .nsmap[["p"]])` to xml2 functions.

### OOXML attribute values

Use the correct string values from the spec, not abbreviations.
Examples: - `<a:bodyPr wrap="square">` not `"sq"` -
`<c:majorTickMark val="out">` not `"outside"` (use
`XL_TICK_MARK$OUTSIDE`) -
`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">`
not `presentationml`

## Test Structure

Tests mirror the R/ file structure in `tests/testthat/`. Each test file
uses `describe()`/`it()` blocks:

``` r
describe("TextFrame", {
  it("sets word_wrap", {
    # ...
  })
})
```

`tests/testthat/helper-xml.R` provides XML-level test helpers.
Integration tests in `test-integration.R` do full round-trips: create →
save → reload → verify.

## Enumerations

All enums are named lists, not factors or integers. Access members as
`MSO_AUTO_SHAPE_TYPE$RECTANGLE`, `XL_CHART_TYPE$COLUMN_CLUSTERED`, etc.
Do not pass raw strings where enums are expected — the OOXML values are
often abbreviated differently from the enum name (e.g.,
`XL_TICK_MARK$OUTSIDE` = `"out"`).
