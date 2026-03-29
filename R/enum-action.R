# Action-related enumeration constants.
#
# Ported from python-pptx/src/pptx/enum/action.py.


#' PP_ACTION_TYPE — type of action for a click or hover
#'
#' Integer constants identifying the type of action associated with a mouse
#' click or hover on a shape or text run.
#'
#' @export
PP_ACTION_TYPE <- list(
  END_SHOW         =   6L,
  FIRST_SLIDE      =   3L,
  HYPERLINK        =   7L,
  LAST_SLIDE       =   4L,
  LAST_SLIDE_VIEWED =  5L,
  NAMED_SLIDE      = 101L,
  NAMED_SLIDE_SHOW =  10L,
  NEXT_SLIDE       =   1L,
  NONE             =   0L,
  OPEN_FILE        = 102L,
  OLE_VERB         =  11L,
  PLAY             =  12L,
  PREVIOUS_SLIDE   =   2L,
  RUN_MACRO        =   8L,
  RUN_PROGRAM      =   9L
)
