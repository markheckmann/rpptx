# Chart data model objects.
#
# Ported from python-pptx/src/pptx/chart/data.py.
# Provides CategoryChartData, XyChartData, BubbleChartData and related
# category/series/data-point classes used to specify chart contents.


# ============================================================================
# Category / Categories
# ============================================================================

#' A single chart category label
#' @noRd
#' @export
Category <- R6::R6Class(
  "Category",

  public = list(
    initialize = function(label, parent) {
      private$.label      <- label
      private$.parent     <- parent
      private$.sub_cats   <- list()
    },

    add_sub_category = function(label) {
      cat <- Category$new(label, self)
      private$.sub_cats <- c(private$.sub_cats, list(cat))
      invisible(cat)
    },

    # Return the 0-based offset of this category in the overall leaf sequence
    index = function(sub_cat) {
      idx <- private$.parent$index(self)
      for (sc in private$.sub_cats) {
        if (identical(sc, sub_cat)) return(idx)
        idx <- idx + sc$leaf_count
      }
      stop("sub_category not in this category", call. = FALSE)
    },

    numeric_str_val = function(date_1904 = FALSE) {
      lbl <- private$.label
      if (inherits(lbl, c("Date", "POSIXt"))) {
        return(sprintf("%.1f", .excel_date_number(lbl, date_1904)))
      }
      as.character(lbl)
    }
  ),

  active = list(
    label = function() {
      lbl <- private$.label
      if (is.null(lbl)) "" else lbl
    },

    # 0-based position in the overall leaf sequence
    idx = function() private$.parent$index(self),

    depth = function() {
      subs <- private$.sub_cats
      if (length(subs) == 0L) return(1L)
      first_depth <- subs[[1]]$depth
      for (s in subs[-1]) {
        if (s$depth != first_depth) stop("category depth not uniform", call. = FALSE)
      }
      first_depth + 1L
    },

    leaf_count = function() {
      if (length(private$.sub_cats) == 0L) return(1L)
      sum(vapply(private$.sub_cats, function(s) s$leaf_count, integer(1L)))
    },

    sub_categories = function() private$.sub_cats
  ),

  private = list(
    .label   = NULL,
    .parent  = NULL,
    .sub_cats = NULL
  )
)


# Convert a Date/POSIXt to Excel serial day number
.excel_date_number <- function(dt, date_1904 = FALSE) {
  d <- as.Date(dt)
  epoch <- if (date_1904) as.Date("1904-01-01") else as.Date("1899-12-31")
  n <- as.integer(d - epoch)
  # Adjust for Excel's false leap year 1900
  if (!date_1904 && n > 59L) n <- n + 1L
  n
}


#' Ordered collection of category labels for a chart
#' @noRd
#' @export
Categories <- R6::R6Class(
  "Categories",

  public = list(
    initialize = function() {
      private$.cats          <- list()
      private$.number_format <- NULL
    },

    add_category = function(label) {
      cat <- Category$new(label, self)
      private$.cats <- c(private$.cats, list(cat))
      invisible(cat)
    },

    # 0-based offset of cat in the leaf sequence
    index = function(cat) {
      idx <- 0L
      for (c in private$.cats) {
        if (identical(c, cat)) return(idx)
        idx <- idx + c$leaf_count
      }
      stop("category not in top-level categories", call. = FALSE)
    },

    # [[i]] accessor (1-based)
    get = function(i) private$.cats[[i]]
  ),

  active = list(
    depth = function() {
      if (length(private$.cats) == 0L) return(0L)
      d <- private$.cats[[1]]$depth
      for (c in private$.cats[-1]) {
        if (c$depth != d) stop("category depth not uniform", call. = FALSE)
      }
      d
    },

    leaf_count = function() {
      sum(vapply(private$.cats, function(c) c$leaf_count, integer(1L)))
    },

    are_dates = function() {
      if (self$depth != 1L) return(FALSE)
      if (length(private$.cats) == 0L) return(FALSE)
      inherits(private$.cats[[1]]$label, c("Date", "POSIXt"))
    },

    are_numeric = function() {
      if (self$depth != 1L) return(FALSE)
      if (length(private$.cats) == 0L) return(FALSE)
      lbl <- private$.cats[[1]]$label
      is.numeric(lbl) || inherits(lbl, c("Date", "POSIXt"))
    },

    number_format = function(value) {
      if (!missing(value)) { private$.number_format <- value; return(invisible(value)) }
      if (!is.null(private$.number_format)) return(private$.number_format)
      if (self$depth != 1L) return("General")
      if (length(private$.cats) == 0L) return("General")
      if (inherits(private$.cats[[1]]$label, c("Date", "POSIXt"))) return("yyyy\\-mm\\-dd")
      "General"
    },

    # Returns a list of levels, each level being a list of list(off, name) pairs.
    # For depth=1 (simple categories): one level, off = 0-based row index.
    levels = function() {
      depth <- self$depth
      if (depth == 0L) return(list())

      # Build level entries recursively
      .levels <- function(cats) {
        sub_cats <- unlist(lapply(cats, function(c) c$sub_categories), recursive = FALSE)
        result <- list()
        if (length(sub_cats) > 0L) {
          result <- c(result, .levels(sub_cats))
        }
        # This level
        this_level <- lapply(cats, function(c) list(off = c$idx, name = c$label))
        result <- c(result, list(this_level))
        result
      }

      .levels(private$.cats)
    }
  ),

  private = list(
    .cats          = NULL,
    .number_format = NULL
  )
)

#' @export
length.Categories <- function(x) length(x$.__enclos_env__$private$.cats)


# ============================================================================
# Base classes
# ============================================================================

#' Base class for chart data objects
#' @noRd
#' @export
BaseChartData <- R6::R6Class(
  "BaseChartData",

  public = list(
    initialize = function(number_format = "General") {
      private$.number_format <- number_format
      private$.series        <- list()
    },

    get = function(i) private$.series[[i]],

    to_list = function() private$.series,

    data_point_offset = function(series) {
      count <- 0L
      for (s in private$.series) {
        if (identical(s, series)) return(count)
        count <- count + s$n_points
      }
      stop("series not in chart data object", call. = FALSE)
    },

    series_index = function(series) {
      for (i in seq_along(private$.series)) {
        if (identical(private$.series[[i]], series)) return(i - 1L)
      }
      stop("series not in chart data object", call. = FALSE)
    },

    series_name_ref = function(series) {
      self$.workbook_writer()$series_name_ref(series)
    },

    x_values_ref = function(series) {
      self$.workbook_writer()$x_values_ref(series)
    },

    y_values_ref = function(series) {
      self$.workbook_writer()$y_values_ref(series)
    },

    xml_str = function(chart_type) {
      chart_xml_writer(chart_type, self)$xml
    },

    # Internal — override per subclass
    .workbook_writer = function() stop("must be implemented by subclass", call. = FALSE)
  ),

  active = list(
    number_format = function() private$.number_format,

    xlsx_blob = function() {
      self$.workbook_writer()$xlsx_blob
    }
  ),

  private = list(
    .number_format = NULL,
    .series        = NULL
  )
)

#' @export
length.BaseChartData <- function(x) length(x$.__enclos_env__$private$.series)


#' Base class for series data objects
#' @noRd
#' @export
BaseSeriesData <- R6::R6Class(
  "BaseSeriesData",

  public = list(
    initialize = function(chart_data, name, number_format = NULL) {
      private$.chart_data    <- chart_data
      private$.name          <- name
      private$.number_format <- number_format
      private$.data_points   <- list()
    }
  ),

  active = list(
    name = function() {
      n <- private$.name
      if (is.null(n)) "" else n
    },

    index = function() private$.chart_data$series_index(self),

    n_points = function() length(private$.data_points),

    data_point_offset = function() private$.chart_data$data_point_offset(self),

    number_format = function() {
      nf <- private$.number_format
      if (is.null(nf)) private$.chart_data$number_format else nf
    },

    name_ref = function() private$.chart_data$series_name_ref(self),

    x_values = function() vapply(private$.data_points, function(dp) dp$x, numeric(1L)),
    y_values = function() vapply(private$.data_points, function(dp) dp$y, numeric(1L))
  ),

  private = list(
    .chart_data    = NULL,
    .name          = NULL,
    .number_format = NULL,
    .data_points   = list()
  )
)


# ============================================================================
# CategoryChartData
# ============================================================================

#' Chart data container for category (bar, line, pie, etc.) charts
#'
#' @description
#' Holds categories and one or more series of values. Use `add_series()` to
#' add data series. Set `categories` by assigning a character/numeric/Date
#' vector to `$categories`.
#'
#' @examples
#' cd <- CategoryChartData$new()
#' cd$categories <- c("Q1", "Q2", "Q3")
#' cd$add_series("Sales", c(100, 200, 150))
#'
#' @noRd
#' @export
CategoryChartData <- R6::R6Class(
  "CategoryChartData",
  inherit = BaseChartData,

  public = list(
    add_series = function(name, values = numeric(0), number_format = NULL) {
      s <- CategorySeriesData$new(self, name, number_format)
      for (v in values) s$add_data_point(v)
      private$.series <- c(private$.series, list(s))
      invisible(s)
    },

    add_category = function(label) {
      self$categories$add_category(label)
    },

    categories_ref = function() self$.workbook_writer()$categories_ref(),

    values_ref = function(series) self$.workbook_writer()$values_ref(series),

    .workbook_writer = function() CategoryWorkbookWriter$new(self)
  ),

  active = list(
    categories = function(value) {
      if (!missing(value)) {
        cats <- Categories$new()
        for (lbl in value) cats$add_category(lbl)
        private$.categories <- cats
        return(invisible(value))
      }
      if (is.null(private$.categories)) private$.categories <- Categories$new()
      private$.categories
    }
  ),

  private = list(
    .categories = NULL
  )
)

#' @noRd
#' @export
ChartData <- CategoryChartData


#' Series data for a category chart
#' @noRd
#' @export
CategorySeriesData <- R6::R6Class(
  "CategorySeriesData",
  inherit = BaseSeriesData,

  public = list(
    add_data_point = function(value, number_format = NULL) {
      dp <- CategoryDataPoint$new(self, value, number_format)
      private$.data_points <- c(private$.data_points, list(dp))
      invisible(dp)
    }
  ),

  active = list(
    categories     = function() private$.chart_data$categories,
    categories_ref = function() private$.chart_data$categories_ref(),
    values_ref     = function() private$.chart_data$values_ref(self),

    values = function() {
      vapply(private$.data_points, function(dp) {
        v <- dp$value
        if (is.null(v)) NA_real_ else as.numeric(v)
      }, numeric(1L))
    }
  )
)

#' @noRd
CategoryDataPoint <- R6::R6Class(
  "CategoryDataPoint",

  public = list(
    initialize = function(series_data, value, number_format = NULL) {
      private$.series_data   <- series_data
      private$.value         <- value
      private$.number_format <- number_format
    }
  ),

  active = list(
    value = function() private$.value,
    number_format = function() {
      nf <- private$.number_format
      if (is.null(nf)) private$.series_data$number_format else nf
    }
  ),

  private = list(
    .series_data   = NULL,
    .value         = NULL,
    .number_format = NULL
  )
)


# ============================================================================
# XyChartData
# ============================================================================

#' Chart data container for XY (scatter) charts
#' @noRd
#' @export
XyChartData <- R6::R6Class(
  "XyChartData",
  inherit = BaseChartData,

  public = list(
    # Add a series. Optionally supply x_values and y_values vectors for convenience.
    add_series = function(name, x_values = NULL, y_values = NULL, number_format = NULL) {
      s <- XySeriesData$new(self, name, number_format)
      private$.series <- c(private$.series, list(s))
      if (!is.null(x_values) && !is.null(y_values)) {
        for (i in seq_along(x_values)) {
          s$add_data_point(x_values[[i]], y_values[[i]])
        }
      }
      invisible(s)
    },

    .workbook_writer = function() XyWorkbookWriter$new(self)
  )
)

#' Series data for an XY chart
#' @noRd
#' @export
XySeriesData <- R6::R6Class(
  "XySeriesData",
  inherit = BaseSeriesData,

  public = list(
    add_data_point = function(x, y, number_format = NULL) {
      dp <- XyDataPoint$new(self, x, y, number_format)
      private$.data_points <- c(private$.data_points, list(dp))
      invisible(dp)
    }
  ),

  active = list(
    x_values_ref = function() private$.chart_data$x_values_ref(self),
    y_values_ref = function() private$.chart_data$y_values_ref(self)
  )
)

#' @noRd
XyDataPoint <- R6::R6Class(
  "XyDataPoint",
  public = list(
    initialize = function(series_data, x, y, number_format = NULL) {
      private$.series_data   <- series_data
      private$.x             <- x
      private$.y             <- y
      private$.number_format <- number_format
    }
  ),
  active = list(
    x = function() private$.x,
    y = function() private$.y,
    number_format = function() {
      nf <- private$.number_format
      if (is.null(nf)) private$.series_data$number_format else nf
    }
  ),
  private = list(
    .series_data   = NULL,
    .x             = NULL,
    .y             = NULL,
    .number_format = NULL
  )
)


# ============================================================================
# BubbleChartData
# ============================================================================

#' Chart data container for bubble charts
#' @noRd
#' @export
BubbleChartData <- R6::R6Class(
  "BubbleChartData",
  inherit = XyChartData,

  public = list(
    add_series = function(name, number_format = NULL) {
      s <- BubbleSeriesData$new(self, name, number_format)
      private$.series <- c(private$.series, list(s))
      invisible(s)
    },

    bubble_sizes_ref = function(series) self$.workbook_writer()$bubble_sizes_ref(series),

    .workbook_writer = function() BubbleWorkbookWriter$new(self)
  )
)

#' Series data for a bubble chart
#' @noRd
#' @export
BubbleSeriesData <- R6::R6Class(
  "BubbleSeriesData",
  inherit = XySeriesData,

  public = list(
    add_data_point = function(x, y, size, number_format = NULL) {
      dp <- BubbleDataPoint$new(self, x, y, size, number_format)
      private$.data_points <- c(private$.data_points, list(dp))
      invisible(dp)
    }
  ),

  active = list(
    bubble_sizes = function() {
      vapply(private$.data_points, function(dp) dp$bubble_size, numeric(1L))
    },
    bubble_sizes_ref = function() private$.chart_data$bubble_sizes_ref(self)
  )
)

#' @noRd
BubbleDataPoint <- R6::R6Class(
  "BubbleDataPoint",
  inherit = XyDataPoint,

  public = list(
    initialize = function(series_data, x, y, size, number_format = NULL) {
      super$initialize(series_data, x, y, number_format)
      private$.size <- size
    }
  ),

  active = list(
    bubble_size = function() private$.size
  ),

  private = list(.size = NULL)
)
