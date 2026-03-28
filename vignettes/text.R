## ----include = FALSE----------------------------------------------------------
knitr::opts_chunk$set(collapse = TRUE, comment = "#>")

## ----setup--------------------------------------------------------------------
library(rpptx)

## ----helper, include=FALSE----------------------------------------------------
shape_with_text <- function() {
  prs    <- Presentation()
  layout <- prs$slide_layouts[[6]]
  slide  <- prs$slides$add_slide(layout)
  slide$shapes$add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(5))
}

