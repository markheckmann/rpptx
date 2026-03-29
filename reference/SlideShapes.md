# Shape collection for a slide

Supports indexed access (1-based),
[`length()`](https://rdrr.io/r/base/length.html), and iteration via
`to_list()`.

## Super classes

[`rpptx::ElementProxy`](https://markheckmann.github.io/rpptx/reference/ElementProxy.md)
-\>
[`rpptx::ParentedElementProxy`](https://markheckmann.github.io/rpptx/reference/ParentedElementProxy.md)
-\> `SlideShapes`

## Methods

### Public methods

- [`SlideShapes$new()`](#method-SlideShapes-new)

- [`SlideShapes$get()`](#method-SlideShapes-get)

- [`SlideShapes$to_list()`](#method-SlideShapes-to_list)

- [`SlideShapes$add_textbox()`](#method-SlideShapes-add_textbox)

- [`SlideShapes$add_shape()`](#method-SlideShapes-add_shape)

- [`SlideShapes$add_table()`](#method-SlideShapes-add_table)

- [`SlideShapes$add_chart()`](#method-SlideShapes-add_chart)

- [`SlideShapes$add_connector()`](#method-SlideShapes-add_connector)

- [`SlideShapes$add_picture()`](#method-SlideShapes-add_picture)

- [`SlideShapes$add_group_shape()`](#method-SlideShapes-add_group_shape)

- [`SlideShapes$build_freeform()`](#method-SlideShapes-build_freeform)

- [`SlideShapes$add_movie()`](#method-SlideShapes-add_movie)

- [`SlideShapes$add_shape_copy()`](#method-SlideShapes-add_shape_copy)

- [`SlideShapes$clone_layout_placeholders()`](#method-SlideShapes-clone_layout_placeholders)

- [`SlideShapes$clone()`](#method-SlideShapes-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    SlideShapes$new(spTree, parent)

------------------------------------------------------------------------

### Method [`get()`](https://rdrr.io/r/base/get.html)

#### Usage

    SlideShapes$get(idx)

------------------------------------------------------------------------

### Method `to_list()`

#### Usage

    SlideShapes$to_list()

------------------------------------------------------------------------

### Method `add_textbox()`

#### Usage

    SlideShapes$add_textbox(left, top, width, height)

------------------------------------------------------------------------

### Method `add_shape()`

#### Usage

    SlideShapes$add_shape(auto_shape_type, left, top, width, height)

------------------------------------------------------------------------

### Method `add_table()`

#### Usage

    SlideShapes$add_table(rows, cols, left, top, width, height)

------------------------------------------------------------------------

### Method `add_chart()`

#### Usage

    SlideShapes$add_chart(chart_type, left, top, width, height, chart_data)

------------------------------------------------------------------------

### Method `add_connector()`

#### Usage

    SlideShapes$add_connector(connector_type, begin_x, begin_y, end_x, end_y)

------------------------------------------------------------------------

### Method `add_picture()`

#### Usage

    SlideShapes$add_picture(image_file, left, top, width = NULL, height = NULL)

------------------------------------------------------------------------

### Method `add_group_shape()`

#### Usage

    SlideShapes$add_group_shape(shapes_list)

------------------------------------------------------------------------

### Method `build_freeform()`

#### Usage

    SlideShapes$build_freeform(start_x = 0, start_y = 0, scale = 1)

------------------------------------------------------------------------

### Method `add_movie()`

#### Usage

    SlideShapes$add_movie(
      movie_file,
      left,
      top,
      width,
      height,
      poster_frame_image = NULL,
      mime_type = "video/unknown"
    )

------------------------------------------------------------------------

### Method `add_shape_copy()`

#### Usage

    SlideShapes$add_shape_copy(shape)

------------------------------------------------------------------------

### Method `clone_layout_placeholders()`

#### Usage

    SlideShapes$clone_layout_placeholders(slide_layout)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    SlideShapes$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
