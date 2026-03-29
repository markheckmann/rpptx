# Action and hyperlink domain objects.
#
# Ported from python-pptx/src/pptx/action.py.
# Provides ActionSetting (click/hover actions on shapes) and ShapeHyperlink.


# ============================================================================
# ShapeHyperlink — hyperlink action on a shape (via cNvPr/hlinkClick)
# ============================================================================

#' Hyperlink on a shape
#'
#' Provides read/write access to a URL hyperlink on a shape.
#' Access via `shape$click_action$hyperlink`.
#'
#' @keywords internal
#' @export
ShapeHyperlink <- R6::R6Class(
  "ShapeHyperlink",

  public = list(
    initialize = function(cNvPr, slide_part, hover = FALSE) {
      private$.cNvPr      <- cNvPr
      private$.slide_part <- slide_part
      private$.hover      <- hover
    }
  ),

  active = list(
    # URL of the hyperlink, or NULL if no hyperlink is set.
    # Assign a string to set a URL; assign NULL to remove.
    address = function(value) {
      if (!missing(value)) {
        private$.remove_hlink()
        if (!is.null(value) && nchar(value) > 0) {
          rId <- private$.slide_part$relate_to(
            value, RT$HYPERLINK, is_external = TRUE
          )
          hlink <- private$.get_or_add_hlink()
          hlink$rId <- rId
        }
        return(invisible(value))
      }
      # Read: look up URL via relationship
      hlink <- private$.hlink()
      if (is.null(hlink)) return(NULL)
      rId <- hlink$rId
      if (is.null(rId)) return(NULL)
      tryCatch({
        rel <- private$.slide_part$rels$get(rId)
        rel$target_ref
      }, error = function(e) NULL)
    }
  ),

  private = list(
    .cNvPr      = NULL,
    .slide_part = NULL,
    .hover      = FALSE,

    .hlink = function() {
      if (private$.hover) return(private$.cNvPr$hlinkHover)
      private$.cNvPr$hlinkClick
    },

    .get_or_add_hlink = function() {
      if (private$.hover) return(private$.cNvPr$get_or_add_hlinkHover())
      private$.cNvPr$get_or_add_hlinkClick()
    },

    .remove_hlink = function() {
      hlink <- private$.hlink()
      if (is.null(hlink)) return(invisible(NULL))
      rId <- hlink$rId
      if (!is.null(rId)) {
        tryCatch(private$.slide_part$drop_rel(rId), error = function(e) NULL)
      }
      if (private$.hover) {
        private$.cNvPr$.remove_hlinkHover()
      } else {
        private$.cNvPr$.remove_hlinkClick()
      }
    }
  )
)


# ============================================================================
# ActionSetting — click/hover action on a shape
# ============================================================================

#' Mouse action settings on a shape
#'
#' Provides access to click and hover action properties of a shape, including
#' hyperlink URL and action type. Access via `shape$click_action`.
#'
#' @keywords internal
#' @export
ActionSetting <- R6::R6Class(
  "ActionSetting",

  public = list(
    initialize = function(cNvPr, parent, hover = FALSE) {
      private$.cNvPr  <- cNvPr
      private$.parent <- parent   # the BaseShape
      private$.hover  <- hover
    }
  ),

  active = list(
    # PP_ACTION_TYPE integer constant indicating the action type.
    action = function() {
      hlink <- private$.hlink()
      if (is.null(hlink)) return(PP_ACTION_TYPE$NONE)

      verb <- hlink$action_verb()

      if (!is.null(verb) && verb == "hlinkshowjump") {
        fields <- hlink$action_fields()
        jump <- fields[["jump"]]
        return(switch(jump,
          "firstslide"      = PP_ACTION_TYPE$FIRST_SLIDE,
          "lastslide"       = PP_ACTION_TYPE$LAST_SLIDE,
          "lastslideviewed" = PP_ACTION_TYPE$LAST_SLIDE_VIEWED,
          "nextslide"       = PP_ACTION_TYPE$NEXT_SLIDE,
          "previousslide"   = PP_ACTION_TYPE$PREVIOUS_SLIDE,
          "endshow"         = PP_ACTION_TYPE$END_SHOW,
          PP_ACTION_TYPE$NONE
        ))
      }

      switch(
        if (is.null(verb)) "hyperlink" else verb,
        "hyperlink"    = PP_ACTION_TYPE$HYPERLINK,
        "hlinksldjump" = PP_ACTION_TYPE$NAMED_SLIDE,
        "hlinkpres"    = PP_ACTION_TYPE$PLAY,
        "hlinkfile"    = PP_ACTION_TYPE$OPEN_FILE,
        "customshow"   = PP_ACTION_TYPE$NAMED_SLIDE_SHOW,
        "ole"          = PP_ACTION_TYPE$OLE_VERB,
        "macro"        = PP_ACTION_TYPE$RUN_MACRO,
        "program"      = PP_ACTION_TYPE$RUN_PROGRAM,
        PP_ACTION_TYPE$NONE
      )
    },

    # ShapeHyperlink for reading/writing the hyperlink URL.
    hyperlink = function(value) {
      if (!missing(value)) return(invisible(NULL))
      ShapeHyperlink$new(private$.cNvPr, private$.parent$part, private$.hover)
    },

    # The target Slide if this is a slide-jump action, else NULL.
    target_slide = function(value) {
      if (!missing(value)) {
        private$clear_click_action()
        if (!is.null(value)) {
          hlink <- private$.get_or_add_hlink()
          hlink$action <- "ppaction://hlinksldjump"
          hlink$rId    <- private$.parent$part$relate_to(
            value$part, RT$SLIDE
          )
        }
        return(invisible(value))
      }
      slide_jump_actions <- c(
        PP_ACTION_TYPE$FIRST_SLIDE,
        PP_ACTION_TYPE$LAST_SLIDE,
        PP_ACTION_TYPE$NEXT_SLIDE,
        PP_ACTION_TYPE$PREVIOUS_SLIDE,
        PP_ACTION_TYPE$NAMED_SLIDE
      )
      act <- self$action
      if (!(act %in% slide_jump_actions)) return(NULL)

      slides <- private$.parent$part$package$presentation_part$presentation$slides
      n      <- length(slides)

      if (act == PP_ACTION_TYPE$FIRST_SLIDE) return(slides[[1]])
      if (act == PP_ACTION_TYPE$LAST_SLIDE)  return(slides[[n]])
      if (act == PP_ACTION_TYPE$NEXT_SLIDE) {
        idx <- private$.slide_index() + 1L
        if (idx > n) stop("no next slide", call. = FALSE)
        return(slides[[idx]])
      }
      if (act == PP_ACTION_TYPE$PREVIOUS_SLIDE) {
        idx <- private$.slide_index() - 1L
        if (idx < 1L) stop("no previous slide", call. = FALSE)
        return(slides[[idx]])
      }
      if (act == PP_ACTION_TYPE$NAMED_SLIDE) {
        hlink     <- private$.hlink()
        rId       <- hlink$rId
        slide_prt <- private$.parent$part$related_part(rId)
        return(slide_prt$slide)
      }
      NULL
    }
  ),

  private = list(
    .cNvPr  = NULL,
    .parent = NULL,
    .hover  = FALSE,

    .hlink = function() {
      if (private$.hover) return(private$.cNvPr$hlinkHover)
      private$.cNvPr$hlinkClick
    },

    .get_or_add_hlink = function() {
      if (private$.hover) return(private$.cNvPr$get_or_add_hlinkHover())
      private$.cNvPr$get_or_add_hlinkClick()
    },

    .slide_index = function() {
      slide_part <- private$.parent$part
      prs_part   <- slide_part$package$presentation_part
      prs        <- prs_part$presentation
      slides     <- prs$slides
      for (i in seq_len(length(slides))) {
        if (identical(slides[[i]]$part, slide_part)) return(i)
      }
      stop("slide not found in presentation", call. = FALSE)
    },

    clear_click_action = function() {
      hlink <- private$.hlink()
      if (is.null(hlink)) return(invisible(NULL))
      rId <- hlink$rId
      if (!is.null(rId)) {
        tryCatch(private$.parent$part$drop_rel(rId), error = function(e) NULL)
      }
      if (private$.hover) {
        private$.cNvPr$.remove_hlinkHover()
      } else {
        private$.cNvPr$.remove_hlinkClick()
      }
    }
  )
)
