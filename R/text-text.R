# Text domain objects: Font, Run, Paragraph, TextFrame.
#
# Ported from python-pptx/src/pptx/text/text.py.

# ============================================================================
# Font — character properties (a:rPr / a:defRPr / a:endParaRPr)
# ============================================================================

#' Character properties for a run, paragraph-default, or end-of-paragraph.
#'
#' Wraps an `a:rPr`, `a:defRPr`, or `a:endParaRPr` element. Provides access
#' to font name, size, bold, italic, and underline. All properties are R/W;
#' assigning NULL removes the override and inherits from the style hierarchy.
#'
#' @include oxml-text.R
#' @keywords internal
#' @export
Font <- R6::R6Class(
  "Font",

  public = list(
    # rPr: CT_TextCharacterProperties element.
    # run: optional parent Run object for effective-value walk-up.
    initialize = function(rPr, run = NULL) {
      private$.rPr <- rPr
      private$.run <- run
    },

    # Underlying CT_TextCharacterProperties element.
    element = function() private$.rPr
  ),

  active = list(
    # Bold: TRUE / FALSE / NULL (inherit).
    bold = function(value) {
      if (!missing(value)) { private$.rPr$b <- value; return(invisible(value)) }
      private$.rPr$b
    },

    # Italic: TRUE / FALSE / NULL (inherit).
    italic = function(value) {
      if (!missing(value)) { private$.rPr$i <- value; return(invisible(value)) }
      private$.rPr$i
    },

    # Font name string or NULL (inherit theme font).
    name = function(value) {
      if (!missing(value)) {
        if (is.null(value)) {
          private$.rPr$`_remove_latin`()
        } else {
          latin <- private$.rPr$get_or_add_latin()
          latin$typeface <- value
        }
        return(invisible(value))
      }
      latin <- private$.rPr$latin
      if (is.null(latin)) return(NULL)
      latin$typeface
    },

    # Font size as a Length (EMU) value, or NULL (inherit).
    # Assign using Pt(n) for point-based values.
    size = function(value) {
      if (!missing(value)) {
        if (is.null(value)) {
          private$.rPr$sz <- NULL
        } else {
          private$.rPr$sz <- as_centipoints(Emu(as.integer(value)))
        }
        return(invisible(value))
      }
      sz <- private$.rPr$sz
      if (is.null(sz)) return(NULL)
      Centipoints(sz)
    },

    # Underline: TRUE (single), FALSE (none), NULL (inherit), or string enum.
    underline = function(value) {
      if (!missing(value)) {
        private$.rPr$u <- if (isTRUE(value)) "sng" else if (identical(value, FALSE)) "none" else value
        return(invisible(value))
      }
      u <- private$.rPr$u
      if (is.null(u)) return(NULL)
      if (u == "none") return(FALSE)
      if (u == "sng")  return(TRUE)
      u
    },

    # ColorFormat for this font's solid fill color. No-op setter for R6 write-back.
    color = function(value) {
      if (!missing(value)) return(invisible(NULL))
      solidFill <- private$.rPr$get_or_add_solidFill()
      ColorFormat$new(solidFill)
    },

    # ---- Effective (inherited) read-only properties -------------------------
    # These cascade: run rPr → paragraph defRPr → txBody lstStyle →
    # layout placeholder lstStyle → master placeholder lstStyle.

    # Effective bold value resolving the inheritance chain. NULL = not set anywhere.
    effective_bold = function() private$.effective_bool("b"),

    # Effective italic value resolving the inheritance chain.
    effective_italic = function() private$.effective_bool("i"),

    # Effective font size resolving the inheritance chain.
    effective_size = function() {
      sz <- private$.rPr$sz
      if (!is.null(sz)) return(Centipoints(sz))
      val <- private$.effective_attr("sz")
      if (is.null(val)) return(NULL)
      Centipoints(as.integer(val))
    },

    # Effective font name resolving the inheritance chain.
    effective_name = function() {
      latin <- private$.rPr$latin
      if (!is.null(latin)) return(latin$typeface)
      private$.effective_latin_typeface()
    }
  ),

  private = list(
    .rPr = NULL,
    .run = NULL,

    # Walk the inheritance chain for a simple boolean-or-NULL attribute.
    .effective_bool = function(attr_name) {
      val <- private$.rPr[[attr_name]]
      if (!is.null(val)) return(val)
      private$.effective_attr(attr_name)
    },

    # Walk the inheritance chain for a raw attribute value (returned as string).
    # Returns NULL if not found anywhere.
    .effective_attr = function(attr_name) {
      if (is.null(private$.run)) return(NULL)
      para <- tryCatch(private$.run$.__enclos_env__$private$.parent,
                       error = function(e) NULL)
      if (is.null(para)) return(NULL)

      # 1. Paragraph's defRPr
      p_elm <- tryCatch(para$.__enclos_env__$private$.p, error = function(e) NULL)
      if (!is.null(p_elm)) {
        pPr <- p_elm$pPr
        if (!is.null(pPr)) {
          defRPr <- pPr$defRPr
          if (!is.null(defRPr)) {
            val <- defRPr[[attr_name]]
            if (!is.null(val)) return(val)
          }
        }
      }

      # 2. txBody lstStyle (paragraph level 1)
      tf <- tryCatch(para$.__enclos_env__$private$.parent, error = function(e) NULL)
      if (!is.null(tf)) {
        txBody <- tryCatch(tf$.__enclos_env__$private$.txBody, error = function(e) NULL)
        if (!is.null(txBody)) {
          val <- private$.attr_from_lstStyle(txBody, 1L, attr_name)
          if (!is.null(val)) return(val)
        }

        # 3. Layout placeholder lstStyle
        shape <- tryCatch(tf$.__enclos_env__$private$.parent, error = function(e) NULL)
        if (!is.null(shape)) {
          layout_ph <- tryCatch(
            shape$.__enclos_env__$private$.base_placeholder(),
            error = function(e) NULL
          )
          if (!is.null(layout_ph)) {
            lph_tf <- tryCatch(layout_ph$text_frame, error = function(e) NULL)
            if (!is.null(lph_tf)) {
              lph_txBody <- tryCatch(
                lph_tf$.__enclos_env__$private$.txBody,
                error = function(e) NULL
              )
              if (!is.null(lph_txBody)) {
                val <- private$.attr_from_lstStyle(lph_txBody, 1L, attr_name)
                if (!is.null(val)) return(val)
              }
            }

            # 4. Master placeholder lstStyle
            master_ph <- tryCatch(
              layout_ph$.__enclos_env__$private$.base_placeholder(),
              error = function(e) NULL
            )
            if (!is.null(master_ph)) {
              mph_tf <- tryCatch(master_ph$text_frame, error = function(e) NULL)
              if (!is.null(mph_tf)) {
                mph_txBody <- tryCatch(
                  mph_tf$.__enclos_env__$private$.txBody,
                  error = function(e) NULL
                )
                if (!is.null(mph_txBody)) {
                  val <- private$.attr_from_lstStyle(mph_txBody, 1L, attr_name)
                  if (!is.null(val)) return(val)
                }
              }
            }
          }
        }
      }
      NULL
    },

    # Extract attr_name from lstStyle/lvl{level}pPr/defRPr.
    .attr_from_lstStyle = function(txBody, level, attr_name) {
      lstStyle <- tryCatch(txBody$lstStyle, error = function(e) NULL)
      if (is.null(lstStyle)) return(NULL)
      lvl_tag <- sprintf("a:lvl%dpPr", level)
      lvlPr_nd <- xml2::xml_find_first(
        lstStyle$get_node(), lvl_tag, ns = c(a = .nsmap[["a"]])
      )
      if (inherits(lvlPr_nd, "xml_missing")) return(NULL)
      defRPr_nd <- xml2::xml_find_first(
        lvlPr_nd, "a:defRPr", ns = c(a = .nsmap[["a"]])
      )
      if (inherits(defRPr_nd, "xml_missing")) return(NULL)
      val <- xml2::xml_attr(defRPr_nd, attr_name)
      if (is.na(val)) return(NULL)
      # For boolean attributes (b, i, strike) coerce to logical
      if (attr_name %in% c("b", "i", "strike", "u")) {
        if (val == "1" || val == "true")  return(TRUE)
        if (val == "0" || val == "false") return(FALSE)
      }
      val
    },

    # Effective latin typeface, walking lstStyle latin elements.
    .effective_latin_typeface = function() {
      if (is.null(private$.run)) return(NULL)
      para <- tryCatch(private$.run$.__enclos_env__$private$.parent,
                       error = function(e) NULL)
      if (is.null(para)) return(NULL)
      tf <- tryCatch(para$.__enclos_env__$private$.parent, error = function(e) NULL)
      if (is.null(tf)) return(NULL)
      txBody <- tryCatch(tf$.__enclos_env__$private$.txBody, error = function(e) NULL)
      if (!is.null(txBody)) {
        val <- private$.latin_from_lstStyle(txBody, 1L)
        if (!is.null(val)) return(val)
      }
      shape <- tryCatch(tf$.__enclos_env__$private$.parent, error = function(e) NULL)
      if (is.null(shape)) return(NULL)
      layout_ph <- tryCatch(shape$.__enclos_env__$private$.base_placeholder(),
                            error = function(e) NULL)
      if (!is.null(layout_ph)) {
        lph_tf <- tryCatch(layout_ph$text_frame, error = function(e) NULL)
        if (!is.null(lph_tf)) {
          lph_body <- tryCatch(lph_tf$.__enclos_env__$private$.txBody,
                               error = function(e) NULL)
          if (!is.null(lph_body)) {
            val <- private$.latin_from_lstStyle(lph_body, 1L)
            if (!is.null(val)) return(val)
          }
        }
        master_ph <- tryCatch(layout_ph$.__enclos_env__$private$.base_placeholder(),
                              error = function(e) NULL)
        if (!is.null(master_ph)) {
          mph_tf <- tryCatch(master_ph$text_frame, error = function(e) NULL)
          if (!is.null(mph_tf)) {
            mph_body <- tryCatch(mph_tf$.__enclos_env__$private$.txBody,
                                 error = function(e) NULL)
            if (!is.null(mph_body)) {
              val <- private$.latin_from_lstStyle(mph_body, 1L)
              if (!is.null(val)) return(val)
            }
          }
        }
      }
      NULL
    },

    .latin_from_lstStyle = function(txBody, level) {
      lstStyle <- tryCatch(txBody$lstStyle, error = function(e) NULL)
      if (is.null(lstStyle)) return(NULL)
      lvl_tag <- sprintf("a:lvl%dpPr", level)
      lvlPr_nd <- xml2::xml_find_first(
        lstStyle$get_node(), lvl_tag, ns = c(a = .nsmap[["a"]])
      )
      if (inherits(lvlPr_nd, "xml_missing")) return(NULL)
      latin_nd <- xml2::xml_find_first(
        lvlPr_nd, "a:defRPr/a:latin", ns = c(a = .nsmap[["a"]])
      )
      if (inherits(latin_nd, "xml_missing")) return(NULL)
      tf <- xml2::xml_attr(latin_nd, "typeface")
      if (is.na(tf)) NULL else tf
    }
  )
)


# ============================================================================
# Hyperlink — wraps a:hlinkClick on a text run
# ============================================================================

#' Hyperlink on a text run
#'
#' Provides read/write access to the URL of a hyperlink on an `a:r` element.
#' Access via `run$hyperlink`.
#'
#' @keywords internal
#' @export
Hyperlink <- R6::R6Class(
  "Hyperlink",

  public = list(
    initialize = function(rPr, slide_part) {
      private$.rPr        <- rPr
      private$.slide_part <- slide_part
    }
  ),

  active = list(
    # URL address of this hyperlink, or NULL if no hyperlink set.
    # Assign a string to set; assign NULL to remove.
    address = function(value) {
      if (!missing(value)) {
        if (is.null(value)) {
          # Remove hlinkClick element if present
          hlink <- private$.rPr$hlinkClick
          if (!is.null(hlink)) {
            rId <- hlink$rId
            private$.rPr$`_remove_hlinkClick`()
            if (!is.null(rId)) {
              tryCatch(private$.slide_part$drop_rel(rId), error = function(e) NULL)
            }
          }
        } else {
          # Add or update hlinkClick with external relationship
          existing_rId <- NULL
          hlink <- private$.rPr$hlinkClick
          if (!is.null(hlink)) existing_rId <- hlink$rId
          # Create relationship (or reuse if same URL)
          rId <- private$.slide_part$relate_to(value, RT$HYPERLINK, is_external = TRUE)
          # Drop old rel if different URL
          if (!is.null(existing_rId) && existing_rId != rId) {
            tryCatch(private$.slide_part$drop_rel(existing_rId), error = function(e) NULL)
          }
          if (is.null(hlink)) {
            hlink <- private$.rPr$get_or_add_hlinkClick()
          }
          hlink$rId <- rId
        }
        return(invisible(value))
      }
      # Read: look up URL via relationship
      hlink <- private$.rPr$hlinkClick
      if (is.null(hlink)) return(NULL)
      rId <- hlink$rId
      if (is.null(rId)) return(NULL)
      tryCatch({
        rel <- private$.slide_part$rels$get(rId)
        rel$target_ref
      }, error = function(e) NULL)
    }
  ),

  private = list(.rPr = NULL, .slide_part = NULL)
)


# ============================================================================
# Run — text run (a:r)
# ============================================================================

#' Text run object corresponding to an `a:r` element.
#'
#' @keywords internal
#' @export
Run <- R6::R6Class(
  "Run",

  public = list(
    initialize = function(r, parent) {
      private$.r      <- r
      private$.parent <- parent
    },

    # Underlying CT_RegularTextRun element.
    element = function() private$.r
  ),

  active = list(
    # Text content of this run (read/write).
    text = function(value) {
      if (!missing(value)) { private$.r$text <- value; return(invisible(value)) }
      private$.r$text
    },

    # Font for run-level character properties.
    # No-op setter accepts the R6 write-back without error.
    font = function(value) {
      if (!missing(value)) return(invisible(NULL))
      rPr <- private$.r$get_or_add_rPr()
      Font$new(rPr, run = self)
    },

    # Hyperlink on this run. Assign NULL to remove.
    # Access via run$hyperlink$address <- "https://..."
    hyperlink = function(value) {
      if (!missing(value)) return(invisible(NULL))
      rPr <- private$.r$get_or_add_rPr()
      slide_part <- tryCatch(private$.parent$part, error = function(e) NULL)
      Hyperlink$new(rPr, slide_part)
    }
  ),

  private = list(.r = NULL, .parent = NULL)
)


# ============================================================================
# Paragraph — paragraph (a:p)
# ============================================================================

#' Paragraph object corresponding to an `a:p` element.
#'
#' @keywords internal
#' @export
Paragraph <- R6::R6Class(
  "Paragraph",

  public = list(
    initialize = function(p, parent) {
      private$.p      <- p
      private$.parent <- parent
    },

    # Underlying CT_TextParagraph element.
    element = function() private$.p,

    # Append a new Run to this paragraph.
    add_run = function() {
      r <- private$.p$add_r()
      Run$new(r, self)
    },

    # Append a soft line break (<a:br>).
    add_line_break = function() {
      private$.p$add_br()
      invisible(self)
    },

    # Remove all content (runs and breaks); preserve paragraph properties.
    clear = function() {
      children <- private$.p$content_children
      for (ch in children) xml2::xml_remove(ch$get_node())
      invisible(self)
    }
  ),

  active = list(
    # The slide Part (resolved via parent chain)
    part = function() private$.parent$part,

    # Paragraph text (read/write). \v represents soft line breaks.
    text = function(value) {
      if (!missing(value)) {
        self$clear()
        private$.p$append_text(value)
        return(invisible(value))
      }
      private$.p$text
    },

    # List of Run objects for this paragraph.
    runs = function(value) {
      if (!missing(value)) return(invisible(NULL))
      lapply(private$.p$r_lst, function(r) Run$new(r, self))
    },

    # Default font for this paragraph (from a:pPr/a:defRPr).
    font = function(value) {
      if (!missing(value)) return(invisible(NULL))
      pPr    <- private$.p$get_or_add_pPr()
      defRPr <- pPr$get_or_add_defRPr()
      Font$new(defRPr)
    },

    # Horizontal alignment string or NULL (inherit).
    alignment = function(value) {
      if (!missing(value)) {
        pPr <- private$.p$get_or_add_pPr(); pPr$algn <- value
        return(invisible(value))
      }
      pPr <- private$.p$pPr
      if (is.null(pPr)) return(NULL)
      pPr$algn
    },

    # Indentation level 0–8 (read/write).
    level = function(value) {
      if (!missing(value)) {
        pPr <- private$.p$get_or_add_pPr(); pPr$lvl <- as.integer(value)
        return(invisible(value))
      }
      pPr <- private$.p$pPr
      if (is.null(pPr)) return(0L)
      pPr$lvl
    },

    # Line spacing: float (lines), Length (points), or NULL.
    line_spacing = function(value) {
      if (!missing(value)) {
        pPr <- private$.p$get_or_add_pPr(); pPr$line_spacing <- value
        return(invisible(value))
      }
      pPr <- private$.p$pPr
      if (is.null(pPr)) return(NULL)
      pPr$line_spacing
    },

    # Space before in EMU (Length) or NULL.
    space_before = function(value) {
      if (!missing(value)) {
        pPr <- private$.p$get_or_add_pPr(); pPr$space_before <- value
        return(invisible(value))
      }
      pPr <- private$.p$pPr
      if (is.null(pPr)) return(NULL)
      pPr$space_before
    },

    # Space after in EMU (Length) or NULL.
    space_after = function(value) {
      if (!missing(value)) {
        pPr <- private$.p$get_or_add_pPr(); pPr$space_after <- value
        return(invisible(value))
      }
      pPr <- private$.p$pPr
      if (is.null(pPr)) return(NULL)
      pPr$space_after
    }
  ),

  private = list(.p = NULL, .parent = NULL)
)


# ============================================================================
# TextFrame — text body (p:txBody)
# ============================================================================

#' Text frame corresponding to a `p:txBody` element.
#'
#' Contains the text of a shape. Access via `shape$text_frame`.
#'
#' @keywords internal
#' @export
TextFrame <- R6::R6Class(
  "TextFrame",

  public = list(
    initialize = function(txBody, parent) {
      private$.txBody <- txBody
      private$.parent <- parent
    },

    # Underlying CT_TextBody element.
    element = function() private$.txBody,

    # Append a new empty paragraph and return a Paragraph wrapper.
    add_paragraph = function() {
      p <- private$.txBody$add_p()
      Paragraph$new(p, self)
    },

    # Remove all paragraphs except the first; clear first paragraph's content.
    clear = function() {
      ps <- private$.txBody$p_lst
      if (length(ps) > 1L) {
        for (p in ps[seq(2L, length(ps))]) xml2::xml_remove(p$get_node())
      }
      first_p <- private$.txBody$p_lst
      if (length(first_p) >= 1L) {
        for (ch in first_p[[1L]]$content_children) xml2::xml_remove(ch$get_node())
      }
      invisible(self)
    }
  ),

  active = list(
    # The slide Part (resolved via parent chain)
    part = function() private$.parent$part,

    # List of Paragraph objects (one per a:p child).
    paragraphs = function(value) {
      if (!missing(value)) return(invisible(NULL))
      lapply(private$.txBody$p_lst, function(p) Paragraph$new(p, self))
    },

    # All text as a single string; paragraphs joined by "\n".
    text = function(value) {
      if (!missing(value)) {
        private$.txBody$clear_content()
        for (p_text in strsplit(as.character(value), "\n", fixed = TRUE)[[1]]) {
          p <- private$.txBody$add_p()
          p$append_text(p_text)
        }
        return(invisible(value))
      }
      paste0(vapply(self$paragraphs, function(p) p$text, character(1)), collapse = "\n")
    },

    # Auto-size behaviour (MSO_AUTO_SIZE constant).
    # NONE (0): no auto-fit (explicit <a:noAutofit/>).
    # SHAPE_TO_FIT_TEXT (1): shape expands to fit text (<a:spAutoFit/>).
    # TEXT_TO_FIT_SHAPE (2): text shrinks to fit shape (<a:normAutofit/>).
    # NULL: no autofit child present (inherit from style).
    auto_size = function(value) {
      if (!missing(value)) {
        bodyPr <- private$.txBody$bodyPr
        if (is.null(value) || identical(value, MSO_AUTO_SIZE$NONE)) {
          bodyPr$get_or_add_noAutofit()
        } else if (identical(value, MSO_AUTO_SIZE$SHAPE_TO_FIT_TEXT)) {
          bodyPr$get_or_add_spAutoFit()
        } else if (identical(value, MSO_AUTO_SIZE$TEXT_TO_FIT_SHAPE)) {
          bodyPr$get_or_add_normAutofit()
        } else {
          stop("auto_size must be an MSO_AUTO_SIZE constant", call. = FALSE)
        }
        return(invisible(value))
      }
      af <- private$.txBody$bodyPr$autofit
      if (is.null(af))                             return(NULL)
      if (inherits(af, "CT_NoAutofit"))            return(MSO_AUTO_SIZE$NONE)
      if (inherits(af, "CT_ShapeAutoFit"))         return(MSO_AUTO_SIZE$SHAPE_TO_FIT_TEXT)
      if (inherits(af, "CT_NormalAutofit"))        return(MSO_AUTO_SIZE$TEXT_TO_FIT_SHAPE)
      NULL
    },

    # Word wrap: TRUE, FALSE, or NULL (inherit).
    word_wrap = function(value) {
      if (!missing(value)) {
        private$.txBody$bodyPr$wrap <-
          if (isTRUE(value)) "sq" else if (identical(value, FALSE)) "none" else NULL
        return(invisible(value))
      }
      w <- private$.txBody$bodyPr$wrap
      if (is.null(w)) return(NULL)
      if (w == "sq")   return(TRUE)
      if (w == "none") return(FALSE)
      NULL
    },

    # Vertical text anchor string or NULL (inherit).
    vertical_anchor = function(value) {
      if (!missing(value)) {
        private$.txBody$bodyPr$anchor <- value
        return(invisible(value))
      }
      private$.txBody$bodyPr$anchor
    },

    # Text insets in EMU (read/write).
    margin_left = function(value) {
      if (!missing(value)) { private$.txBody$bodyPr$lIns <- value; return(invisible(value)) }
      private$.txBody$bodyPr$lIns
    },
    margin_right = function(value) {
      if (!missing(value)) { private$.txBody$bodyPr$rIns <- value; return(invisible(value)) }
      private$.txBody$bodyPr$rIns
    },
    margin_top = function(value) {
      if (!missing(value)) { private$.txBody$bodyPr$tIns <- value; return(invisible(value)) }
      private$.txBody$bodyPr$tIns
    },
    margin_bottom = function(value) {
      if (!missing(value)) { private$.txBody$bodyPr$bIns <- value; return(invisible(value)) }
      private$.txBody$bodyPr$bIns
    }
  ),

  private = list(.txBody = NULL, .parent = NULL)
)


#' @export
length.TextFrame <- function(x) {
  length(x$.__enclos_env__$private$.txBody$p_lst)
}

#' @export
`[[.TextFrame` <- function(x, i) x$paragraphs[[i]]
