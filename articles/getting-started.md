# Getting Started with rpptx

``` r
library(rpptx)
```

`rpptx` is an R port of the Python `python-pptx` library. It lets you
create, read, and update PowerPoint (`.pptx`) files entirely from R,
with an R6-based object model that closely mirrors the Python API.

## Creating a presentation

[`pptx_presentation()`](https://markheckmann.github.io/rpptx/reference/pptx_presentation.md)
opens an existing file or — with no argument — creates a new blank
presentation based on the built-in default template.

``` r
prs <- pptx_presentation()
```

Check the slide dimensions (in EMU — English Metric Units, 914400 per
inch):

``` r
prs$slide_width   # 10 inches
#> <Length: 9144000 EMU (10.00 in, 25.40 cm, 720.0 pt)>
prs$slide_height  # 7.5 inches
#> <Length: 6858000 EMU (7.50 in, 19.05 cm, 540.0 pt)>
```

Use
[`Inches()`](https://markheckmann.github.io/rpptx/reference/Length.md),
[`Cm()`](https://markheckmann.github.io/rpptx/reference/Length.md), or
[`Pt()`](https://markheckmann.github.io/rpptx/reference/Length.md) to
convert from convenient units to EMU:

``` r
Inches(1)     # 914400L
#> <Length: 914400 EMU (1.00 in, 2.54 cm, 72.0 pt)>
Cm(2.54)      # 914400L
#> <Length: 914400 EMU (1.00 in, 2.54 cm, 72.0 pt)>
Pt(72)        # 914400L
#> <Length: 914400 EMU (1.00 in, 2.54 cm, 72.0 pt)>
```

To convert back from EMU to a readable unit use
[`as_inches()`](https://markheckmann.github.io/rpptx/reference/as_inches.md),
[`as_cm()`](https://markheckmann.github.io/rpptx/reference/as_inches.md),
[`as_mm()`](https://markheckmann.github.io/rpptx/reference/as_inches.md),
or
[`as_pt()`](https://markheckmann.github.io/rpptx/reference/as_inches.md):

``` r
as_inches(prs$slide_width)   # 10
#> [1] 10
as_cm(prs$slide_width)       # 25.4
#> [1] 25.4
as_pt(Pt(18))                # 18
#> [1] 18
```

## Adding slides

Slides are added through the `Slides` collection using a slide layout as
a template. Every new blank presentation has one slide master with
eleven layouts.

``` r
# List available layout names
sapply(seq_len(length(prs$slide_layouts)), function(i) {
  prs$slide_layouts[[i]]$name
})
#>  [1] "Title Slide"             "Title and Content"      
#>  [3] "Section Header"          "Two Content"            
#>  [5] "Comparison"              "Title Only"             
#>  [7] "Blank"                   "Content with Caption"   
#>  [9] "Picture with Caption"    "Title and Vertical Text"
#> [11] "Vertical Title and Text"
```

Add a blank slide (layout index 6 is “Blank” in the default template):

``` r
layout <- prs$slide_layouts[[6]]
slide  <- prs$slides$add_slide(layout)
```

Access individual slides by 1-based index:

``` r
length(prs$slides)    # 1
#> [1] 1
prs$slides[[1]]       # the slide we just added
#> <Slide>  1 shape
#>   Placeholder     Title                     9.00×1.25 in @ 0.50, 0.30
```

## Adding a text box

``` r
slide$shapes$add_textbox(
  left   = Inches(1),
  top    = Inches(1),
  width  = Inches(5),
  height = Inches(1)
)
#> <TextBox>  "TextBox 2"  5.00×1.00 in  @ 1.00, 1.00 in
shape <- slide$shapes[[1]]
shape$text_frame$text <- "Hello, rpptx!"
```

## Adding a rectangle

``` r
rect <- slide$shapes$add_shape(
  MSO_AUTO_SHAPE_TYPE$ROUNDED_RECTANGLE,
  left   = Inches(1),
  top    = Inches(2.5),
  width  = Inches(3),
  height = Inches(1.5)
)
rect$text <- "A rounded rectangle"
rect$fill$solid()
rect$fill$fore_color$rgb <- RGBColor(0x46L, 0x72L, 0xC4L)
```

## Adding a picture

``` r
slide$shapes$add_picture(
  "path/to/photo.jpg",
  left   = Inches(5),
  top    = Inches(2),
  width  = Inches(3),
  height = Inches(2)
)
```

## Saving

``` r
prs$save("my_presentation.pptx")
```

Or save to a temp file:

``` r
tmp <- tempfile(fileext = ".pptx")
prs$save(tmp)
file.exists(tmp)
#> [1] TRUE
```

## Opening an existing file

``` r
prs2 <- pptx_presentation("existing.pptx")
slide1 <- prs2$slides[[1]]
for (shape in slide1$shapes$to_list()) {
  cat(shape$name, "\n")
}
```

## Slide operations

### Deleting a slide

``` r
prs2 <- pptx_presentation()
layout <- prs2$slide_layouts[[6]]
s1 <- prs2$slides$add_slide(layout)
s2 <- prs2$slides$add_slide(layout)
length(prs2$slides)   # 2
#> [1] 2
prs2$slides$delete(s1)
length(prs2$slides)   # 1
#> [1] 1
```

### Reordering slides

``` r
prs3 <- pptx_presentation()
layout <- prs3$slide_layouts[[6]]
for (i in 1:3) prs3$slides$add_slide(layout)
id_last <- prs3$slides[[3]]$slide_id
prs3$slides$move(prs3$slides[[3]], 1L)
prs3$slides[[1]]$slide_id == id_last  # TRUE
#> [1] TRUE
```

## Core properties

``` r
prs4 <- pptx_presentation()
cp <- prs4$core_properties
cp$title   <- "My Presentation"
cp$author  <- "Jane Doe"
cp$subject <- "rpptx demo"
cp$title
#> [1] "My Presentation"
```

## Unit conversion reference

| Function         | Units       | EMU         |
|------------------|-------------|-------------|
| `Emu(n)`         | EMU         | n           |
| `Inches(n)`      | inches      | n × 914 400 |
| `Cm(n)`          | centimetres | n × 360 000 |
| `Pt(n)`          | points      | n × 12 700  |
| `Mm(n)`          | millimetres | n × 36 000  |
| `Centipoints(n)` | centipoints | n × 127     |
