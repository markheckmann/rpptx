# Working with Text

``` r
library(rpptx)
```

## The text object model

Text in rpptx follows a three-level hierarchy mirroring the OOXML
structure:

    TextFrame  (p:txBody)
      └─ Paragraph  (a:p)
           └─ Run  (a:r)

A `TextFrame` belongs to a shape and contains one or more `Paragraph`
objects. Each `Paragraph` contains one or more `Run` objects. Runs are
the smallest unit with independent character formatting.

## Accessing a TextFrame

Any `Shape` (autoshape or text box) with `has_text_frame == TRUE`
exposes `text_frame`:

``` r
shape <- shape_with_text()
tf    <- shape$text_frame
class(tf)  # "TextFrame"
#> [1] "TextFrame" "R6"
```

## Quick text access

For convenience, `shape$text` reads/writes all text as a single string:

``` r
shape <- shape_with_text()
shape$text <- "Hello world"
shape$text
#> [1] "Hello world"
```

`text_frame$text` does the same at the frame level:

``` r
shape$text_frame$text <- "Updated text"
shape$text_frame$text
#> [1] "Updated text"
```

## Paragraphs

``` r
shape <- shape_with_text()
tf    <- shape$text_frame

# Add paragraphs
p1 <- tf$add_paragraph()
p1$text <- "First paragraph"

p2 <- tf$add_paragraph()
p2$text <- "Second paragraph"

length(tf$paragraphs)  # 2 (there's always a default empty paragraph first)
#> [1] 3
```

Paragraph alignment:

``` r
p1$alignment <- PP_PARAGRAPH_ALIGNMENT$CENTER
p2$alignment <- PP_PARAGRAPH_ALIGNMENT$RIGHT
```

Alignment values: `LEFT`, `CENTER`, `RIGHT`, `JUSTIFY`, `DISTRIBUTE`.

Indentation level (0–8):

``` r
p2$level <- 1L   # indent one level
```

Line and paragraph spacing (in points):

``` r
p1$line_spacing  <- 1.5    # 1.5× line height
p1$space_before  <- Pt(6)
p1$space_after   <- Pt(12)
```

## Runs and character formatting

A run carries font formatting. Add runs to a paragraph:

``` r
shape <- shape_with_text()
tf    <- shape$text_frame
tf$clear()

p <- tf$add_paragraph()
r1 <- p$add_run(); r1$text <- "Bold "
r2 <- p$add_run(); r2$text <- "and italic"
```

### Font properties

``` r
r1$font$bold   <- TRUE
r2$font$italic <- TRUE
```

``` r
r1$font$size          <- Pt(18)        # font size
r1$font$name          <- "Calibri"
r1$font$color$rgb     <- RGBColor(0xFF, 0x00, 0x00)  # red
r1$font$underline     <- TRUE
```

Access the paragraph-level default font (applies to all runs without
explicit overrides):

``` r
p$font$bold <- TRUE    # makes all runs in p bold by default
```

### Theme colours

``` r
r2$font$color$theme_color <- MSO_THEME_COLOR$ACCENT_1
```

### Clearing formatting (inherit from theme)

``` r
r1$font$bold  <- NULL   # removes override, inherits from theme
r1$font$size  <- NULL
```

## Hyperlinks

Attach a URL to a run:

``` r
shape <- shape_with_text()
tf    <- shape$text_frame
tf$clear()
p <- tf$add_paragraph()
r <- p$add_run()
r$text <- "Visit our website"
r$hyperlink$address <- "https://example.com"
```

Read back:

``` r
r$hyperlink$address   # "https://example.com"
```

## Text frame options

### Word wrap

``` r
shape <- shape_with_text()
shape$text_frame$word_wrap <- TRUE   # wrap text (default)
shape$text_frame$word_wrap <- FALSE  # no wrap
```

### Vertical anchor

``` r
shape$text_frame$vertical_anchor <- MSO_VERTICAL_ANCHOR$MIDDLE
```

Values: `TOP`, `MIDDLE`, `BOTTOM`.

### Margins (internal padding)

``` r
tf <- shape$text_frame
tf$margin_left   <- Inches(0.1)
tf$margin_right  <- Inches(0.1)
tf$margin_top    <- Inches(0.05)
tf$margin_bottom <- Inches(0.05)
```

## Building rich text from scratch

``` r
prs    <- pptx_presentation()
layout <- prs$slide_layouts[[6]]
slide  <- prs$slides$add_slide(layout)

tb <- slide$shapes$add_textbox(
  Inches(1), Inches(1), Inches(8), Inches(5.5)
)
tf <- tb$text_frame
tf$word_wrap <- TRUE

# First paragraph — heading style
tf$clear()
p1 <- tf$paragraphs[[1]]
p1$text <- "rpptx Text Demo"
p1$font$bold <- TRUE
p1$font$size <- Pt(28)
p1$alignment <- PP_PARAGRAPH_ALIGNMENT$CENTER

# Second paragraph — normal text with mixed formatting
p2 <- tf$add_paragraph()
p2$space_before <- Pt(12)
r1 <- p2$add_run(); r1$text <- "You can mix "
r2 <- p2$add_run(); r2$text <- "bold"; r2$font$bold <- TRUE
r3 <- p2$add_run(); r3$text <- ", "
r4 <- p2$add_run(); r4$text <- "italic"; r4$font$italic <- TRUE
r5 <- p2$add_run(); r5$text <- ", and "
r6 <- p2$add_run()
r6$text <- "coloured"
r6$font$color$rgb <- RGBColor(0xC0, 0x00, 0x00)
r7 <- p2$add_run(); r7$text <- " text freely."

# Third paragraph — bulleted list simulation
bullets <- c("First item", "Second item", "Third item")
for (b in bullets) {
  p <- tf$add_paragraph()
  p$text  <- b
  p$level <- 1L
  p$space_before <- Pt(4)
}

tmp <- tempfile(fileext = ".pptx")
prs$save(tmp)
```

## Notes text

Each slide can have speaker notes accessible via `notes_text_frame`:

``` r
prs    <- pptx_presentation()
layout <- prs$slide_layouts[[6]]
slide  <- prs$slides$add_slide(layout)

ns  <- slide$notes_slide
ntf <- ns$notes_text_frame
ntf$text <- "These are the speaker notes for this slide."
ntf$text
#> [1] "These are the speaker notes for this slide."
```
