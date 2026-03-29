# FreeformBuilder and drawing-operation classes.
#
# Ported from python-pptx/src/pptx/shapes/freeform.py.
# Allows a freeform (custom-geometry) shape to be specified and added to a slide.

#' @include oxml-shapes.R shapes-base.R
#' @keywords internal


# ============================================================================
# FreeformBuilder
# ============================================================================

#' Build a freeform (custom-geometry) shape
#'
#' Obtained via `slide$shapes$build_freeform(...)`. Allows successive calls to
#' `add_line_segments()` and `move_to()` to define the shape geometry, then
#' `convert_to_shape()` to add it to the slide.
#'
#' @keywords internal
#' @export
FreeformBuilder <- R6::R6Class(
  "FreeformBuilder",

  public = list(
    initialize = function(spTree, shapes, start_x, start_y, x_scale, y_scale) {
      private$.spTree  <- spTree
      private$.shapes  <- shapes
      private$.start_x <- Emu(as.integer(start_x))
      private$.start_y <- Emu(as.integer(start_y))
      private$.x_scale <- x_scale
      private$.y_scale <- y_scale
      private$.ops     <- list()
    },

    # ---- drawing API ----

    # Add a straight line segment to each point in vertices.
    # vertices: list of two-element numeric vectors c(x, y) in local coords.
    # close: if TRUE (default) a closing segment is added.
    # Returns self for chaining.
    add_line_segments = function(vertices, close = TRUE) {
      for (v in vertices) {
        private$.ops <- c(private$.ops, list(.LineSegment$new(self, v[[1]], v[[2]])))
      }
      if (close) private$.ops <- c(private$.ops, list(.Close$new()))
      invisible(self)
    },

    # Move the pen to (x, y) without drawing.  Returns self for chaining.
    move_to = function(x, y) {
      private$.ops <- c(private$.ops, list(.MoveTo$new(self, x, y)))
      invisible(self)
    },

    # Convert drawing operations into a freeform shape on the slide.
    # origin_x / origin_y locate the local-coordinate origin in slide EMU.
    # May be called more than once to stamp the same geometry at different positions.
    # Returns the new Shape proxy object.
    convert_to_shape = function(origin_x = Emu(0L), origin_y = Emu(0L)) {
      sp   <- private$.add_freeform_sp(origin_x, origin_y)
      path <- private$.start_path(sp)
      for (op in private$.ops) op$apply_operation_to(path)
      shape_factory(sp, private$.shapes)
    },

    # ---- read-only geometry accessors used by drawing operations ----

    # Return x of leftmost extent in local coordinates (Emu).
    shape_offset_x = function() {
      min_x <- private$.start_x
      for (op in private$.ops) {
        if (!inherits(op, ".Close")) min_x <- min(min_x, op$x)
      }
      Emu(min_x)
    },

    # Return y of topmost extent in local coordinates (Emu).
    shape_offset_y = function() {
      min_y <- private$.start_y
      for (op in private$.ops) {
        if (!inherits(op, ".Close")) min_y <- min(min_y, op$y)
      }
      Emu(min_y)
    }
  ),

  private = list(
    .spTree  = NULL,   # CT_GroupShape element (spTree)
    .shapes  = NULL,   # SlideShapes collection (for shape_factory)
    .start_x = NULL,
    .start_y = NULL,
    .x_scale = NULL,
    .y_scale = NULL,
    .ops     = NULL,   # list of drawing-operation objects

    # ---- internal geometry helpers ----

    .dx = function() {
      min_x <- max_x <- private$.start_x
      for (op in private$.ops) {
        if (!inherits(op, ".Close")) {
          min_x <- min(min_x, op$x)
          max_x <- max(max_x, op$x)
        }
      }
      Emu(max_x - min_x)
    },

    .dy = function() {
      min_y <- max_y <- private$.start_y
      for (op in private$.ops) {
        if (!inherits(op, ".Close")) {
          min_y <- min(min_y, op$y)
          max_y <- max(max_y, op$y)
        }
      }
      Emu(max_y - min_y)
    },

    .width = function() {
      as.integer(round(private$.dx() * private$.x_scale))
    },

    .height = function() {
      as.integer(round(private$.dy() * private$.y_scale))
    },

    .left = function() {
      as.integer(round(self$shape_offset_x() * private$.x_scale))
    },

    .top = function() {
      as.integer(round(self$shape_offset_y() * private$.y_scale))
    },

    # Add freeform sp element to the slide spTree and return it
    .add_freeform_sp = function(origin_x, origin_y) {
      x  <- as.integer(origin_x) + private$.left()
      y  <- as.integer(origin_y) + private$.top()
      cx <- private$.width()
      cy <- private$.height()
      private$.spTree$add_freeform_sp(x, y, cx, cy)
    },

    # Create the initial <a:path> inside sp and add the start moveTo
    .start_path = function(sp) {
      path <- sp$add_path(w = private$.dx(), h = private$.dy())
      sx <- Emu(as.integer(private$.start_x) - as.integer(self$shape_offset_x()))
      sy <- Emu(as.integer(private$.start_y) - as.integer(self$shape_offset_y()))
      path$add_moveTo(sx, sy)
      path
    }
  )
)


# ============================================================================
# Drawing-operation helper classes (internal)
# ============================================================================

# Base class ----
.BaseDrawingOperation <- R6::R6Class(
  ".BaseDrawingOperation",
  public = list(
    initialize = function(builder, x, y) {
      private$.builder <- builder
      private$.x <- Emu(as.integer(round(x)))
      private$.y <- Emu(as.integer(round(y)))
    },
    apply_operation_to = function(path) stop("not implemented", call. = FALSE)
  ),
  private = list(
    .builder = NULL,
    .x       = NULL,
    .y       = NULL
  ),
  active = list(
    x = function() private$.x,
    y = function() private$.y
  )
)


# Close ----
.Close <- R6::R6Class(
  ".Close",
  public = list(
    apply_operation_to = function(path) path$add_close()
  )
)


# LineSegment ----
.LineSegment <- R6::R6Class(
  ".LineSegment",
  inherit = .BaseDrawingOperation,
  public = list(
    apply_operation_to = function(path) {
      path$add_lnTo(
        Emu(as.integer(private$.x) - as.integer(private$.builder$shape_offset_x())),
        Emu(as.integer(private$.y) - as.integer(private$.builder$shape_offset_y()))
      )
    }
  )
)


# MoveTo ----
.MoveTo <- R6::R6Class(
  ".MoveTo",
  inherit = .BaseDrawingOperation,
  public = list(
    apply_operation_to = function(path) {
      path$add_moveTo(
        Emu(as.integer(private$.x) - as.integer(private$.builder$shape_offset_x())),
        Emu(as.integer(private$.y) - as.integer(private$.builder$shape_offset_y()))
      )
    }
  )
)
