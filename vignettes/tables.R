## ----include = FALSE----------------------------------------------------------
knitr::opts_chunk$set(collapse = TRUE, comment = "#>")

## ----setup--------------------------------------------------------------------
library(rpptx)

## ----helper, include=FALSE----------------------------------------------------
blank_slide <- function() {
  prs    <- Presentation()
  layout <- prs$slide_layouts[[6]]
  list(prs = prs, slide = prs$slides$add_slide(layout))
}

