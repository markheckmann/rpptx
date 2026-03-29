## ----include = FALSE----------------------------------------------------------
knitr::opts_chunk$set(collapse = TRUE, comment = "#>")

## ----setup--------------------------------------------------------------------
library(rpptx)

## ----helper, include=FALSE----------------------------------------------------
blank_slide <- function() {
  prs    <- pptx_presentation()
  layout <- prs$slide_layouts[[6]]
  list(prs = prs, slide = prs$slides$add_slide(layout))
}

new_rect <- function(slide) {
  slide$shapes$add_shape(
    MSO_AUTO_SHAPE_TYPE$RECTANGLE,
    Inches(1), Inches(1), Inches(3), Inches(2)
  )
}

## -----------------------------------------------------------------------------
red   <- RGBColor(255L, 0L, 0L)
green <- RGBColor(0L, 128L, 0L)
blue  <- RGBColor_from_str("0000FF")   # from hex string

print(red)     # RGBColor(255, 0, 0) [#FF0000]
as.character(blue)  # "0000FF"
red == RGBColor(255L, 0L, 0L)  # TRUE

## -----------------------------------------------------------------------------
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$solid()
shape$fill$fore_color$rgb <- RGBColor(0x4F, 0x81, 0xBD)
shape$fill$type   # "solid"

## -----------------------------------------------------------------------------
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$solid()
shape$fill$fore_color$theme_color <- MSO_THEME_COLOR$ACCENT_1
shape$fill$fore_color$type   # "scheme"

## -----------------------------------------------------------------------------
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$solid()
shape$fill$fore_color$theme_color <- MSO_THEME_COLOR$ACCENT_1
shape$fill$fore_color$tint  <- 0.4   # 40 % lighter
shape$fill$fore_color$shade <- NULL  # remove shade (if any)

## -----------------------------------------------------------------------------
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)
shape$fill$background()
shape$fill$type   # "background"

## -----------------------------------------------------------------------------
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$gradient()
shape$fill$gradient_angle <- 270.0   # top-to-bottom

stops <- shape$fill$gradient_stops
stops[[1]]$color$rgb <- RGBColor(0x4F, 0x81, 0xBD)  # start color
stops[[2]]$color$rgb <- RGBColor(0xBD, 0xD7, 0xEE)  # end color

length(stops)               # 2
stops[[1]]$position         # 0
stops[[2]]$position         # 1

## -----------------------------------------------------------------------------
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$fill$patterned()
shape$fill$pattern <- MSO_PATTERN_TYPE$DIAGONAL_BRICK
shape$fill$fore_color$rgb <- RGBColor(0x4F, 0x81, 0xBD)
shape$fill$back_color$rgb <- RGBColor(0xFF, 0xFF, 0xFF)

shape$fill$type     # "patterned"
shape$fill$pattern  # "diagBrick"

## -----------------------------------------------------------------------------
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)

shape$line$width     <- Pt(2)                         # 2-point line
shape$line$color$rgb <- RGBColor(0xFF, 0, 0)          # red
shape$line$dash_style <- MSO_LINE_DASH_STYLE$DASH      # dashed

shape$line$width      # 182880 EMU ≈ 2 pt

## -----------------------------------------------------------------------------
shape$line$width <- NULL   # inherit from theme

## -----------------------------------------------------------------------------
r <- blank_slide(); slide <- r$slide
shape <- new_rect(slide)
shape$fill$solid()

shadow <- shape$shadow
shadow$inherit        # TRUE (inherits from theme)
shadow$visible        # FALSE by default unless theme enables it

## -----------------------------------------------------------------------------
r <- blank_slide(); slide <- r$slide
slide$background$fill$solid()
slide$background$fill$fore_color$rgb <- RGBColor(0x1F, 0x1F, 0x1F)

