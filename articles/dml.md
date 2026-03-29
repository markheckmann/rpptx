# Colors, Fills, and Lines

``` r
library(rpptx)
```

## RGB colors

`RGBColor` holds a red–green–blue triple (0–255 each):

``` r
red   <- RGBColor(255L, 0L, 0L)
green <- RGBColor(0L, 128L, 0L)
blue  <- RGBColor_from_str("0000FF")   # from hex string

print(red)     # RGBColor(255, 0, 0) [#FF0000]
#> RGBColor(255, 0, 0) [#FF0000]
as.character(blue)  # "0000FF"
#> [1] "0000FF"
red == RGBColor(255L, 0L, 0L)  # TRUE
#> [1] TRUE
```

## Solid fills

Call `fill$solid()` then set the foreground color:

``` r
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$solid()
shape$fill$fore_color$rgb <- RGBColor(0x4F, 0x81, 0xBD)
shape$fill$type   # "solid"
#> [1] "solid"
```

### Theme (scheme) colors

Use an `MSO_THEME_COLOR` value to apply a color that follows the
presentation theme:

``` r
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$solid()
shape$fill$fore_color$theme_color <- MSO_THEME_COLOR$ACCENT_1
shape$fill$fore_color$type   # "scheme"
#> [1] "scheme"
```

Apply a **tint** (lighten toward white) or **shade** (darken toward
black):

``` r
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$solid()
shape$fill$fore_color$theme_color <- MSO_THEME_COLOR$ACCENT_1
shape$fill$fore_color$tint  <- 0.4   # 40 % lighter
shape$fill$fore_color$shade <- NULL  # remove shade (if any)
```

## No fill (transparent)

``` r
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)
shape$fill$background()
shape$fill$type   # "background"
#> NULL
```

## Gradient fills

`fill$gradient()` creates a default two-stop linear gradient. Customise
the stops and angle afterwards:

``` r
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$gradient()
shape$fill$gradient_angle <- 270.0   # top-to-bottom

stops <- shape$fill$gradient_stops
stops[[1]]$color$rgb <- RGBColor(0x4F, 0x81, 0xBD)  # start color
stops[[2]]$color$rgb <- RGBColor(0xBD, 0xD7, 0xEE)  # end color

length(stops)               # 2
#> [1] 2
stops[[1]]$position         # 0
#> [1] 0
stops[[2]]$position         # 1
#> [1] 1
```

## Pattern fills

`fill$patterned()` creates a hatched pattern. Set the preset with
`fill$pattern` and control foreground/background colors with
`fore_color`/`back_color`:

``` r
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$patterned()
shape$fill$pattern <- MSO_PATTERN_TYPE$DIAGONAL_BRICK
shape$fill$fore_color$rgb <- RGBColor(0x4F, 0x81, 0xBD)
shape$fill$back_color$rgb <- RGBColor(0xFF, 0xFF, 0xFF)

shape$fill$type     # "patterned"
#> [1] "patterned"
shape$fill$pattern  # "diagBrick"
#> [1] "diagBrick"
```

## Line formatting

Access line properties via `shape$line`:

``` r
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$line$width     <- Pt(2)                         # 2-point line
shape$line$color$rgb <- RGBColor(0xFF, 0, 0)          # red
shape$line$dash_style <- MSO_LINE_DASH_STYLE$DASH      # dashed

shape$line$width      # 182880 EMU ≈ 2 pt
#> <Length: 25400 EMU (0.03 in, 0.07 cm, 2.0 pt)>
```

Set to `NULL` to remove the width override:

``` r
shape$line$width <- NULL   # inherit from theme
```

## Shadow effects

`shape$shadow` exposes a `ShadowFormat` object:

``` r
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)
shape$fill$solid()

shadow <- shape$shadow
shadow$inherit        # TRUE (inherits from theme)
#> [1] TRUE
shadow$visible        # FALSE by default unless theme enables it
#> NULL
```

## Slide background

Each slide exposes a `background` object with a `fill` property.
Accessing `background$fill` interrupts master-level background
inheritance and applies an explicit fill to this slide:

``` r
r <- blank_slide(); slide <- r$slide
slide$background$fill$solid()
slide$background$fill$fore_color$rgb <- RGBColor(0x1F, 0x1F, 0x1F)
```

## Putting it together: a styled shape

``` r
r <- blank_slide()
prs   <- r$prs
slide <- r$slide

shape <- slide$shapes$add_shape(
  MSO_AUTO_SHAPE_TYPE$ROUNDED_RECTANGLE,
  Inches(1), Inches(1), Inches(4), Inches(2)
)

# Gradient fill: blue to sky-blue
shape$fill$gradient()
shape$fill$gradient_angle <- 135.0
shape$fill$gradient_stops[[1]]$color$rgb <- RGBColor(0x00, 0x70, 0xC0)
shape$fill$gradient_stops[[2]]$color$rgb <- RGBColor(0x9D, 0xC3, 0xE6)

# White border
shape$line$width     <- Pt(1.5)
shape$line$color$rgb <- RGBColor(0xFF, 0xFF, 0xFF)

# Text
tf  <- shape$text_frame
run <- tf$paragraphs[[1]]$add_run()
run$text <- "Styled shape"
tf$paragraphs[[1]]$runs[[1]]$font$color$rgb <- RGBColor(0xFF, 0xFF, 0xFF)
tf$paragraphs[[1]]$runs[[1]]$font$bold      <- TRUE

tmp <- tempfile(fileext = ".pptx")
prs$save(tmp)
```
