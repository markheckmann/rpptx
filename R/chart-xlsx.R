# Workbook writers for chart data (embedded Excel files).
#
# Ported from python-pptx/src/pptx/chart/xlsx.py.
# Uses openxlsx2 (listed in Imports) instead of Python's xlsxwriter.


# Convert 1-based column number to Excel column letter notation.
# E.g. 1 -> "A", 26 -> "Z", 27 -> "AA", 703 -> "AAA".
.excel_col_letter <- function(col_number) {
  if (col_number < 1L || col_number > 16384L)
    stop("column_number must be in range 1-16384", call. = FALSE)
  col_ref <- ""
  n <- col_number
  while (n > 0L) {
    remainder <- n %% 26L
    if (remainder == 0L) remainder <- 26L
    col_ref <- paste0(intToUtf8(utf8ToInt("A") + remainder - 1L), col_ref)
    n <- (n - 1L) %/% 26L
  }
  col_ref
}


# ============================================================================
# .BaseWorkbookWriter — shared base
# ============================================================================

.BaseWorkbookWriter <- R6::R6Class(
  ".BaseWorkbookWriter",

  public = list(
    initialize = function(chart_data) {
      private$.chart_data <- chart_data
    }
  ),

  active = list(
    xlsx_blob = function() {
      if (!requireNamespace("openxlsx2", quietly = TRUE))
        stop(paste0(
          "openxlsx2 is required for chart support. ",
          "Install it with: install.packages('openxlsx2')"
        ), call. = FALSE)

      wb <- openxlsx2::wb_workbook()
      wb$add_worksheet("Sheet1")
      private$.populate_worksheet(wb)

      tmp <- tempfile(fileext = ".xlsx")
      on.exit(unlink(tmp), add = TRUE)
      wb$save(tmp)
      readBin(tmp, raw(), n = file.size(tmp))
    }
  ),

  private = list(
    .chart_data = NULL,

    # Write data into wb$Sheet1; override in subclasses.
    .populate_worksheet = function(wb) {
      stop("must be implemented by subclass", call. = FALSE)
    },

    # Write a column of values into wb at (start_row, start_col), 1-based.
    .write_col = function(wb, values, start_row, start_col) {
      if (length(values) == 0L) return(invisible(NULL))
      for (i in seq_along(values)) {
        wb$add_data(
          sheet    = 1L,
          x        = values[[i]],
          start_row = start_row + i - 1L,
          start_col = start_col,
          colNames  = FALSE
        )
      }
    }
  )
)


# ============================================================================
# CategoryWorkbookWriter
# ============================================================================

#' Workbook writer for category (bar, column, line, pie, etc.) charts
#' @keywords internal
#' @export
CategoryWorkbookWriter <- R6::R6Class(
  "CategoryWorkbookWriter",
  inherit = .BaseWorkbookWriter,

  public = list(
    # Excel cell-range reference to the categories (excluding header row).
    categories_ref = function() {
      cats <- private$.chart_data$categories
      if (cats$depth == 0L)
        stop("chart data contains no categories", call. = FALSE)
      right_col   <- .excel_col_letter(cats$depth)
      bottom_row  <- cats$leaf_count + 1L
      sprintf("Sheet1!$A$2:$%s$%d", right_col, bottom_row)
    },

    # Excel cell reference to the header cell (series name) for `series`.
    series_name_ref = function(series) {
      sprintf("Sheet1!$%s$1", private$.series_col_letter(series))
    },

    # Excel range reference to the values for `series`.
    values_ref = function(series) {
      col_letter <- private$.series_col_letter(series)
      bottom_row <- series$n_points + 1L
      sprintf("Sheet1!$%s$2:$%s$%d", col_letter, col_letter, bottom_row)
    }
  ),

  private = list(
    # Letter of the column containing data for `series`.
    .series_col_letter = function(series) {
      col_num <- 1L + series$categories$depth + series$index
      .excel_col_letter(col_num)
    },

    .populate_worksheet = function(wb) {
      private$.write_categories(wb)
      private$.write_series(wb)
    },

    # Write category label column(s). Each level occupies one column, and
    # the deepest level is in column A (depth-1=0 in 0-based, col 1 in 1-based).
    .write_categories = function(wb) {
      cats  <- private$.chart_data$categories
      depth <- cats$depth
      lvls  <- cats$levels
      for (idx in seq_along(lvls)) {
        level <- lvls[[idx]]
        col   <- depth - idx + 1L  # 1-based; innermost level -> col 1
        for (entry in level) {
          row <- entry$off + 2L  # 0-based offset + header row + 1-based
          wb$add_data(
            sheet     = 1L,
            x         = entry$name,
            start_row = row,
            start_col = col,
            colNames  = FALSE
          )
        }
      }
    },

    # Write series name (row 1) and values (rows 2..n+1) for each series.
    .write_series = function(wb) {
      col_offset   <- private$.chart_data$categories$depth + 1L
      series_list  <- private$.chart_data$to_list()
      for (i in seq_along(series_list)) {
        series     <- series_list[[i]]
        series_col <- col_offset + i - 1L
        wb$add_data(
          sheet     = 1L,
          x         = series$name,
          start_row = 1L,
          start_col = series_col,
          colNames  = FALSE
        )
        private$.write_col(wb, series$values, 2L, series_col)
      }
    }
  )
)


# ============================================================================
# XyWorkbookWriter
# ============================================================================

#' Workbook writer for XY (scatter) charts
#' @keywords internal
#' @export
XyWorkbookWriter <- R6::R6Class(
  "XyWorkbookWriter",
  inherit = .BaseWorkbookWriter,

  public = list(
    # Excel cell reference for the series name (header of Y column).
    series_name_ref = function(series) {
      row <- private$.series_table_row_offset(series) + 1L
      sprintf("Sheet1!$B$%d", row)
    },

    # Excel range reference to the X values for `series`.
    x_values_ref = function(series) {
      top_row    <- private$.series_table_row_offset(series) + 2L
      bottom_row <- top_row + series$n_points - 1L
      sprintf("Sheet1!$A$%d:$A$%d", top_row, bottom_row)
    },

    # Excel range reference to the Y values for `series`.
    y_values_ref = function(series) {
      top_row    <- private$.series_table_row_offset(series) + 2L
      bottom_row <- top_row + series$n_points - 1L
      sprintf("Sheet1!$B$%d:$B$%d", top_row, bottom_row)
    }
  ),

  private = list(
    # 0-based row offset for the title row of this series' data table.
    .series_table_row_offset = function(series) {
      title_and_spacer_rows <- series$index * 2L
      data_point_rows       <- series$data_point_offset
      title_and_spacer_rows + data_point_rows
    },

    .populate_worksheet = function(wb) {
      series_list <- private$.chart_data$to_list()
      for (series in series_list) {
        offset <- private$.series_table_row_offset(series)
        # X values in col A, starting at row offset+2 (1-based)
        private$.write_col(wb, series$x_values, offset + 2L, 1L)
        # Series name in col B, row offset+1 (1-based)
        wb$add_data(
          sheet     = 1L,
          x         = series$name,
          start_row = offset + 1L,
          start_col = 2L,
          colNames  = FALSE
        )
        # Y values in col B, starting at row offset+2 (1-based)
        private$.write_col(wb, series$y_values, offset + 2L, 2L)
      }
    }
  )
)


# ============================================================================
# BubbleWorkbookWriter
# ============================================================================

#' Workbook writer for bubble charts
#' @keywords internal
#' @export
BubbleWorkbookWriter <- R6::R6Class(
  "BubbleWorkbookWriter",
  inherit = XyWorkbookWriter,

  public = list(
    # Excel range reference to the bubble sizes for `series`.
    bubble_sizes_ref = function(series) {
      top_row    <- private$.series_table_row_offset(series) + 2L
      bottom_row <- top_row + series$n_points - 1L
      sprintf("Sheet1!$C$%d:$C$%d", top_row, bottom_row)
    }
  ),

  private = list(
    .populate_worksheet = function(wb) {
      series_list <- private$.chart_data$to_list()
      for (series in series_list) {
        offset <- private$.series_table_row_offset(series)
        private$.write_col(wb, series$x_values, offset + 2L, 1L)
        wb$add_data(
          sheet     = 1L,
          x         = series$name,
          start_row = offset + 1L,
          start_col = 2L,
          colNames  = FALSE
        )
        private$.write_col(wb, series$y_values, offset + 2L, 2L)
        wb$add_data(
          sheet     = 1L,
          x         = "Size",
          start_row = offset + 1L,
          start_col = 3L,
          colNames  = FALSE
        )
        private$.write_col(wb, series$bubble_sizes, offset + 2L, 3L)
      }
    }
  )
)
