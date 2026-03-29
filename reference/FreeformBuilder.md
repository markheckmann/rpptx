# Freeform shape builder

Specifies and creates a freeform (custom-geometry) shape. Obtain via
`slide$shapes$build_freeform()`.

## Methods

- `add_line_segments(vertices, close = TRUE)`:

  Add line segments from current position to each point in `vertices` (a
  list of `c(x, y)` vectors in local coordinates). If `close` is `TRUE`
  (default) a closing segment is appended. Returns `self` invisibly for
  chaining.

- `move_to(x, y)`:

  Move pen to `(x, y)` without drawing. Returns `self` invisibly for
  chaining.

- `convert_to_shape(origin_x = Emu(0), origin_y = Emu(0))`:

  Create the freeform shape on the slide. `origin_x` and `origin_y` (in
  EMU) locate the local-coordinate origin on the slide. Returns the new
  `Shape` proxy object.

- `shape_offset_x()`:

  Return the x coordinate of the leftmost extent (in local units).

- `shape_offset_y()`:

  Return the y coordinate of the topmost extent (in local units).

## Examples

``` r
if (FALSE) { # \dontrun{
prs   <- pptx_presentation()
slide <- prs$slides$add_slide(prs$slide_layouts[[6]])

# local coords where 100 units = 1 inch
scale <- Inches(1) / 100

ff <- slide$shapes$build_freeform(start_x = 0, start_y = 0, scale = scale)
ff$add_line_segments(list(c(100, 0), c(50, 87), c(0, 0)))  # triangle
shp <- ff$convert_to_shape(origin_x = Inches(1), origin_y = Inches(1))
prs$save("output.pptx")
} # }
```
