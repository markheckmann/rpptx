# Chart domain objects — Chart, Axis, Legend, Series, Plot, etc.
#
# Ported from python-pptx/src/pptx/chart/


# ============================================================================
# ChartFormat — access to fill and line for chart elements
# ============================================================================

#' Chart element format (fill + line)
#'
#' Provides `fill` and `line` properties for chart sub-elements like axes,
#' series, markers, and gridlines.
#'
#' @noRd
ChartFormat <- R6::R6Class(
  "ChartFormat",
  public = list(
    initialize = function(element) {
      private$.element <- element
    }
  ),

  active = list(
    # FillFormat for this element's shape properties
    fill = function(value) {
      if (!missing(value)) return(invisible(NULL))
      spPr <- private$.element$get_or_add_spPr()
      FillFormat$from_fill_parent(spPr)
    },
    # LineFormat for this element's shape properties
    line = function(value) {
      if (!missing(value)) return(invisible(NULL))
      spPr <- private$.element$get_or_add_spPr()
      LineFormat$new(spPr)
    }
  ),

  private = list(.element = NULL)
)


# ============================================================================
# Marker — data point marker on line/XY/radar charts
# ============================================================================

#' Data point marker object
#'
#' Controls size and style of markers on line-type chart series.
#'
#' @noRd
Marker <- R6::R6Class(
  "Marker",
  public = list(
    initialize = function(element) {
      private$.element <- element
    }
  ),

  active = list(
    format = function(value) {
      if (!missing(value)) return(invisible(NULL))
      marker <- private$.element$get_or_add_marker()
      ChartFormat$new(marker)
    },

    size = function(value) {
      marker <- private$.element$marker
      if (!missing(value)) {
        m <- private$.element$get_or_add_marker()
        m$.remove_size()
        if (!is.null(value)) { sz <- m$.add_size(); sz$val <- value }
        return(invisible(value))
      }
      if (is.null(marker)) return(NULL)
      marker$size_val
    },

    style = function(value) {
      marker <- private$.element$marker
      if (!missing(value)) {
        m <- private$.element$get_or_add_marker()
        m$.remove_symbol()
        if (!is.null(value)) { sy <- m$.add_symbol(); sy$val <- value }
        return(invisible(value))
      }
      if (is.null(marker)) return(NULL)
      marker$symbol_val
    }
  ),

  private = list(.element = NULL)
)


# ============================================================================
# DataLabels — data label collection for a plot or series
# ============================================================================

#' Data labels collection object
#'
#' Controls display properties of all data labels in a plot or series.
#'
#' @noRd
DataLabels <- R6::R6Class(
  "DataLabels",
  public = list(
    initialize = function(dLbls) {
      private$.element <- dLbls
    }
  ),

  active = list(
    font = function(value) {
      if (!missing(value)) return(invisible(NULL))
      defRPr <- private$.element$defRPr
      Font$new(defRPr)
    },

    number_format = function(value) {
      if (!missing(value)) {
        nm_elm <- private$.element$get_or_add_numFmt()
        nm_elm$formatCode <- value
        self$number_format_is_linked <- FALSE
        return(invisible(value))
      }
      numFmt <- private$.element$numFmt
      if (is.null(numFmt)) return("General")
      numFmt$formatCode
    },

    number_format_is_linked = function(value) {
      if (!missing(value)) {
        nm_elm <- private$.element$get_or_add_numFmt()
        nm_elm$sourceLinked <- value
        return(invisible(value))
      }
      numFmt <- private$.element$numFmt
      if (is.null(numFmt)) return(TRUE)
      sl <- numFmt$sourceLinked
      if (is.null(sl)) return(TRUE)
      sl
    },

    position = function(value) {
      if (!missing(value)) {
        if (is.null(value)) {
          private$.element$.remove_dLblPos()
        } else {
          pos_elm <- private$.element$get_or_add_dLblPos()
          pos_elm$val <- value
        }
        return(invisible(value))
      }
      dLblPos <- private$.element$dLblPos
      if (is.null(dLblPos)) return(NULL)
      dLblPos$val
    },

    show_category_name = function(value) {
      elm <- private$.element$get_or_add_showCatName()
      if (!missing(value)) { elm$val <- isTRUE(value); return(invisible(value)) }
      elm$val
    },

    show_legend_key = function(value) {
      elm <- private$.element$get_or_add_showLegendKey()
      if (!missing(value)) { elm$val <- isTRUE(value); return(invisible(value)) }
      elm$val
    },

    show_percentage = function(value) {
      elm <- private$.element$get_or_add_showPercent()
      if (!missing(value)) { elm$val <- isTRUE(value); return(invisible(value)) }
      elm$val
    },

    show_series_name = function(value) {
      elm <- private$.element$get_or_add_showSerName()
      if (!missing(value)) { elm$val <- isTRUE(value); return(invisible(value)) }
      elm$val
    },

    show_value = function(value) {
      elm <- private$.element$get_or_add_showVal()
      if (!missing(value)) { elm$val <- isTRUE(value); return(invisible(value)) }
      elm$val
    }
  ),

  private = list(.element = NULL)
)


# ============================================================================
# DataLabel — individual data point label
# ============================================================================

#' Individual data point label object
#' @noRd
DataLabel <- R6::R6Class(
  "DataLabel",
  public = list(
    initialize = function(ser, idx) {
      private$.ser <- ser
      private$.idx <- idx
    },

    .get_or_add_dLbl = function() private$.ser$get_or_add_dLbl(private$.idx),
    .get_or_add_rich = function() {
      dLbl <- self$.get_or_add_dLbl()
      dLbl$.remove_spPr()
      dLbl$.remove_txPr()
      dLbl$get_or_add_rich()
    }
  ),

  active = list(
    font = function(value) {
      if (!missing(value)) return(invisible(NULL))
      dLbl   <- self$.get_or_add_dLbl()
      txPr   <- dLbl$get_or_add_txPr()
      tf     <- TextFrame$new(txPr, self)
      tf$paragraphs[[1]]$font
    },

    has_text_frame = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) {
          self$.get_or_add_rich()
        } else {
          dLbl <- private$.ser$get_dLbl(private$.idx)
          if (!is.null(dLbl)) dLbl$remove_tx_rich()
        }
        return(invisible(value))
      }
      dLbl <- private$.ser$get_dLbl(private$.idx)
      if (is.null(dLbl)) return(FALSE)
      length(dLbl$xpath("c:tx/c:rich")) > 0
    },

    position = function(value) {
      if (!missing(value)) {
        if (is.null(value)) {
          dLbl <- private$.ser$get_dLbl(private$.idx)
          if (!is.null(dLbl)) dLbl$.remove_dLblPos()
        } else {
          { dl <- self$.get_or_add_dLbl(); dp <- dl$get_or_add_dLblPos(); dp$val <- value }
        }
        return(invisible(value))
      }
      dLbl <- private$.ser$get_dLbl(private$.idx)
      if (is.null(dLbl)) return(NULL)
      pos <- dLbl$dLblPos
      if (is.null(pos)) return(NULL)
      pos$val
    },

    text_frame = function(value) {
      if (!missing(value)) return(invisible(NULL))
      rich <- self$.get_or_add_rich()
      TextFrame$new(rich, self)
    }
  ),

  private = list(.ser = NULL, .idx = NULL)
)


# ============================================================================
# Point, CategoryPoints, XyPoints, BubblePoints
# ============================================================================

#' Individual data point in a series
#' @noRd
Point <- R6::R6Class(
  "Point",
  public = list(
    initialize = function(ser, idx) {
      private$.ser <- ser
      private$.idx <- idx
    }
  ),
  active = list(
    data_label = function(value) {
      if (!missing(value)) return(invisible(NULL))
      DataLabel$new(private$.ser, private$.idx)
    },
    format = function(value) {
      if (!missing(value)) return(invisible(NULL))
      dPt <- private$.ser$get_or_add_dPt_for_point(private$.idx)
      ChartFormat$new(dPt)
    },
    marker = function(value) {
      if (!missing(value)) return(invisible(NULL))
      dPt <- private$.ser$get_or_add_dPt_for_point(private$.idx)
      Marker$new(dPt)
    }
  ),
  private = list(.ser = NULL, .idx = NULL)
)

#' Point collection for category-based series
#' @noRd
CategoryPoints <- R6::R6Class(
  "CategoryPoints",
  public = list(
    initialize = function(ser) private$.ser <- ser,
    get = function(idx) {
      n <- private$.ser$cat_ptCount_val
      if (idx < 0L || idx >= n) stop("point index out of range", call. = FALSE)
      Point$new(private$.ser, idx)
    }
  ),
  private = list(.ser = NULL)
)

#' @export
length.CategoryPoints <- function(x) x$.__enclos_env__$private$.ser$cat_ptCount_val

#' @export
`[[.CategoryPoints` <- function(x, i) x$get(i - 1L)  # 1-based to 0-based

#' Point collection for XY (scatter) series
#' @noRd
XyPoints <- R6::R6Class(
  "XyPoints",
  public = list(
    initialize = function(ser) private$.ser <- ser,
    get = function(idx) {
      n <- min(private$.ser$xVal_ptCount_val, private$.ser$yVal_ptCount_val)
      if (idx < 0L || idx >= n) stop("point index out of range", call. = FALSE)
      Point$new(private$.ser, idx)
    }
  ),
  private = list(.ser = NULL)
)

#' @export
length.XyPoints <- function(x) {
  ser <- x$.__enclos_env__$private$.ser
  min(ser$xVal_ptCount_val, ser$yVal_ptCount_val)
}

#' @export
`[[.XyPoints` <- function(x, i) x$get(i - 1L)

#' Point collection for bubble series
#' @noRd
BubblePoints <- R6::R6Class(
  "BubblePoints",
  public = list(
    initialize = function(ser) private$.ser <- ser,
    get = function(idx) {
      ser <- private$.ser
      n <- min(ser$xVal_ptCount_val, ser$yVal_ptCount_val, ser$bubbleSize_ptCount_val)
      if (idx < 0L || idx >= n) stop("point index out of range", call. = FALSE)
      Point$new(ser, idx)
    }
  ),
  private = list(.ser = NULL)
)

#' @export
length.BubblePoints <- function(x) {
  ser <- x$.__enclos_env__$private$.ser
  min(ser$xVal_ptCount_val, ser$yVal_ptCount_val, ser$bubbleSize_ptCount_val)
}

#' @export
`[[.BubblePoints` <- function(x, i) x$get(i - 1L)


# ============================================================================
# Series classes
# ============================================================================

#' Base class for chart series objects
#' @noRd
BaseSeries <- R6::R6Class(
  "BaseSeries",
  public = list(
    initialize = function(ser) {
      private$.element <- ser
      private$.ser     <- ser
    }
  ),
  active = list(
    format = function(value) {
      if (!missing(value)) return(invisible(NULL))
      ChartFormat$new(private$.ser)
    },
    index = function() private$.element$idx$val,
    name = function() {
      nds <- private$.element$xpath("./c:tx//c:pt/c:v")
      if (length(nds) == 0) return("")
      xml2::xml_text(nds[[1]])
    }
  ),
  private = list(.element = NULL, .ser = NULL)
)

#' Base class for category-type series (bar, line, pie, area, radar)
#' @noRd
BaseCategorySeries <- R6::R6Class(
  "BaseCategorySeries",
  inherit = BaseSeries,
  active = list(
    data_labels = function(value) {
      if (!missing(value)) return(invisible(NULL))
      DataLabels$new(private$.ser$get_or_add_dLbls())
    },
    points = function(value) {
      if (!missing(value)) return(invisible(NULL))
      CategoryPoints$new(private$.ser)
    },
    values = function() {
      val_elm <- private$.element$val
      if (is.null(val_elm)) return(numeric(0))
      n <- val_elm$ptCount_val
      vapply(seq_len(n) - 1L, function(i) {
        v <- val_elm$pt_v(i)
        if (is.null(v)) NA_real_ else v
      }, numeric(1))
    }
  )
)

#' Area series
#' @noRd
AreaSeries <- R6::R6Class("AreaSeries", inherit = BaseCategorySeries)

#' Bar series
#' @noRd
BarSeries <- R6::R6Class(
  "BarSeries",
  inherit = BaseCategorySeries,
  active = list(
    invert_if_negative = function(value) {
      ifn <- private$.element$invertIfNegative
      if (!missing(value)) {
        { ifn <- private$.element$get_or_add_invertIfNegative(); ifn$val <- isTRUE(value) }
        return(invisible(value))
      }
      if (is.null(ifn)) return(TRUE)
      ifn$val
    }
  )
)

#' Line series
#' @noRd
LineSeries <- R6::R6Class(
  "LineSeries",
  inherit = BaseCategorySeries,
  active = list(
    marker = function(value) {
      if (!missing(value)) return(invisible(NULL))
      Marker$new(private$.ser)
    },
    smooth = function(value) {
      sm <- private$.element$smooth
      if (!missing(value)) {
        { sm <- private$.element$get_or_add_smooth(); sm$val <- isTRUE(value) }
        return(invisible(value))
      }
      if (is.null(sm)) return(TRUE)
      sm$val
    }
  )
)

#' Pie series
#' @noRd
PieSeries <- R6::R6Class("PieSeries", inherit = BaseCategorySeries)

#' Radar series
#' @noRd
RadarSeries <- R6::R6Class(
  "RadarSeries",
  inherit = BaseCategorySeries,
  active = list(
    marker = function(value) {
      if (!missing(value)) return(invisible(NULL))
      Marker$new(private$.ser)
    }
  )
)

#' XY scatter series
#' @noRd
XySeries <- R6::R6Class(
  "XySeries",
  inherit = BaseSeries,
  active = list(
    marker = function(value) {
      if (!missing(value)) return(invisible(NULL))
      Marker$new(private$.ser)
    },
    points = function(value) {
      if (!missing(value)) return(invisible(NULL))
      XyPoints$new(private$.ser)
    },
    values = function() {
      yVal <- private$.element$yVal
      if (is.null(yVal)) return(numeric(0))
      n <- yVal$ptCount_val
      vapply(seq_len(n) - 1L, function(i) {
        v <- yVal$pt_v(i)
        if (is.null(v)) NA_real_ else v
      }, numeric(1))
    }
  )
)

#' Bubble series
#' @noRd
BubbleSeries <- R6::R6Class(
  "BubbleSeries",
  inherit = XySeries,
  active = list(
    points = function(value) {
      if (!missing(value)) return(invisible(NULL))
      BubblePoints$new(private$.ser)
    }
  )
)


#' Dispatch a c:ser element to the appropriate series class
#'
#' @param ser A CT_SeriesComposite wrapped element.
#' @return A series domain object.
#' @noRd
series_factory <- function(ser) {
  # ser's parent determines the plot type
  parent_nd <- xml2::xml_parent(ser$get_node())
  parent_tag <- xml2::xml_name(parent_nd, ns = xml2::xml_ns(parent_nd))
  cls <- switch(parent_tag,
    "c:areaChart"    = ,
    "c:area3DChart"  = AreaSeries,
    "c:barChart"     = ,
    "c:bar3DChart"   = BarSeries,
    "c:bubbleChart"  = BubbleSeries,
    "c:doughnutChart"= PieSeries,
    "c:lineChart"    = ,
    "c:line3DChart"  = LineSeries,
    "c:pieChart"     = ,
    "c:pie3DChart"   = PieSeries,
    "c:radarChart"   = RadarSeries,
    "c:scatterChart" = XySeries,
    BaseSeries
  )
  cls$new(ser)
}


# ============================================================================
# SeriesCollection
# ============================================================================

#' Collection of series in a chart or plot
#' @noRd
SeriesCollection <- R6::R6Class(
  "SeriesCollection",
  public = list(
    initialize = function(parent_elm) {
      # parent_elm can be CT_PlotArea or CT_xChart
      private$.element <- parent_elm
    },
    get = function(idx) {
      srs <- private$.element$sers
      if (idx < 1L || idx > length(srs)) stop("series index out of range", call. = FALSE)
      series_factory(srs[[idx]])
    },
    to_list = function() lapply(private$.element$sers, series_factory)
  ),
  private = list(.element = NULL)
)

#' @export
length.SeriesCollection <- function(x) length(x$.__enclos_env__$private$.element$sers)

#' @export
`[[.SeriesCollection` <- function(x, i) x$get(i)


# ============================================================================
# Plot classes
# ============================================================================

#' Base class for chart plot objects
#' @noRd
BasePlot <- R6::R6Class(
  "BasePlot",
  public = list(
    initialize = function(xChart, chart) {
      private$.element <- xChart
      private$.chart   <- chart
    }
  ),
  active = list(
    chart = function() private$.chart,

    series = function(value) {
      if (!missing(value)) return(invisible(NULL))
      SeriesCollection$new(private$.element)
    },

    data_labels = function(value) {
      if (!missing(value)) return(invisible(NULL))
      dLbls <- private$.element$dLbls
      if (is.null(dLbls)) stop("plot has no data labels; set has_data_labels = TRUE first",
                               call. = FALSE)
      DataLabels$new(dLbls)
    },

    has_data_labels = function(value) {
      if (!missing(value)) {
        if (!isTRUE(value)) {
          private$.element$.remove_dLbls()
        } else if (is.null(private$.element$dLbls)) {
          dLbls <- private$.element$add_dLbls_default()
          sv_elm <- dLbls$get_or_add_showVal()
          sv_elm$val <- TRUE
        }
        return(invisible(value))
      }
      !is.null(private$.element$dLbls)
    },

    vary_by_categories = function(value) {
      vc <- private$.element$varyColors
      if (!missing(value)) {
        { vc <- private$.element$get_or_add_varyColors(); vc$val <- isTRUE(value) }
        return(invisible(value))
      }
      if (is.null(vc)) return(TRUE)
      vc$val
    }
  ),
  private = list(.element = NULL, .chart = NULL)
)

#' Area chart plot
#' @noRd
AreaPlot <- R6::R6Class("AreaPlot", inherit = BasePlot)

#' 3-D area chart plot
#' @noRd
Area3DPlot <- R6::R6Class("Area3DPlot", inherit = BasePlot)

#' Bar or column chart plot
#' @noRd
BarPlot <- R6::R6Class(
  "BarPlot",
  inherit = BasePlot,
  active = list(
    gap_width = function(value) {
      gw <- private$.element$gapWidth
      if (!missing(value)) {
        { gw <- private$.element$get_or_add_gapWidth(); gw$val <- value }
        return(invisible(value))
      }
      if (is.null(gw)) return(150L)
      gw$val
    },
    overlap = function(value) {
      ov <- private$.element$overlap
      if (!missing(value)) {
        if (value == 0L) {
          private$.element$.remove_overlap()
        } else {
          { ov <- private$.element$get_or_add_overlap(); ov$val <- value }
        }
        return(invisible(value))
      }
      if (is.null(ov)) return(0L)
      ov$val
    }
  )
)

#' Bubble chart plot
#' @noRd
BubblePlot <- R6::R6Class("BubblePlot", inherit = BasePlot)

#' Doughnut chart plot
#' @noRd
DoughnutPlot <- R6::R6Class("DoughnutPlot", inherit = BasePlot)

#' Line chart plot
#' @noRd
LinePlot <- R6::R6Class("LinePlot", inherit = BasePlot)

#' Pie chart plot
#' @noRd
PiePlot <- R6::R6Class("PiePlot", inherit = BasePlot)

#' Radar chart plot
#' @noRd
RadarPlot <- R6::R6Class("RadarPlot", inherit = BasePlot)

#' XY (scatter) chart plot
#' @noRd
XyPlot <- R6::R6Class("XyPlot", inherit = BasePlot)


#' Dispatch an xChart element to the appropriate plot class
#' @param xChart A CT_xChart wrapped element.
#' @param chart The Chart domain object.
#' @return A plot domain object.
#' @noRd
plot_factory <- function(xChart, chart) {
  nd <- xChart$get_node()
  tag <- xml2::xml_name(nd, ns = xml2::xml_ns(nd))
  cls <- switch(tag,
    "c:areaChart"     = AreaPlot,
    "c:area3DChart"   = Area3DPlot,
    "c:barChart"      = ,
    "c:bar3DChart"    = BarPlot,
    "c:bubbleChart"   = BubblePlot,
    "c:doughnutChart" = DoughnutPlot,
    "c:lineChart"     = ,
    "c:line3DChart"   = LinePlot,
    "c:pieChart"      = ,
    "c:pie3DChart"    = PiePlot,
    "c:radarChart"    = RadarPlot,
    "c:scatterChart"  = XyPlot,
    BasePlot
  )
  cls$new(xChart, chart)
}


#' Determine chart type of a plot
#' @noRd
plot_type_inspector <- function(plot) {
  XL <- XL_CHART_TYPE
  nd <- plot$.__enclos_env__$private$.element$get_node()
  tag <- xml2::xml_name(nd, ns = xml2::xml_ns(nd))

  grouping_val <- function() {
    gv_nd <- xml2::xml_find_first(nd, "c:grouping", ns = c(c = .nsmap[["c"]]))
    if (inherits(gv_nd, "xml_missing")) return(NULL)
    xml2::xml_attr(gv_nd, "val")
  }

  switch(class(plot)[1],
    "AreaPlot" = {
      gv <- grouping_val()
      switch(gv,
        "standard"        = XL$AREA,
        "stacked"         = XL$AREA_STACKED,
        "percentStacked"  = XL$AREA_STACKED_100,
        XL$AREA
      )
    },
    "Area3DPlot" = {
      gv <- grouping_val()
      switch(gv,
        "standard"        = XL$THREE_D_AREA,
        "stacked"         = XL$THREE_D_AREA_STACKED,
        "percentStacked"  = XL$THREE_D_AREA_STACKED_100,
        XL$THREE_D_AREA
      )
    },
    "BarPlot" = {
      barDir_nd <- xml2::xml_find_first(nd, "c:barDir", ns = c(c = .nsmap[["c"]]))
      bd <- xml2::xml_attr(barDir_nd, "val")
      gv <- grouping_val()
      if (bd == "bar") {
        switch(gv,
          "clustered"       = XL$BAR_CLUSTERED,
          "stacked"         = XL$BAR_STACKED,
          "percentStacked"  = XL$BAR_STACKED_100,
          XL$BAR_CLUSTERED
        )
      } else {
        switch(gv,
          "clustered"       = XL$COLUMN_CLUSTERED,
          "stacked"         = XL$COLUMN_STACKED,
          "percentStacked"  = XL$COLUMN_STACKED_100,
          XL$COLUMN_CLUSTERED
        )
      }
    },
    "BubblePlot" = {
      b3d <- xml2::xml_find_all(nd, "c:ser/c:bubble3D", ns = c(c = .nsmap[["c"]]))
      if (length(b3d) == 0) return(XL$BUBBLE)
      v_attr <- xml2::xml_attr(b3d[[1]], "val")
      if (!is.na(v_attr) && !(v_attr %in% c("0","false"))) XL$BUBBLE_THREE_D_EFFECT else XL$BUBBLE
    },
    "DoughnutPlot" = {
      exp_nds <- xml2::xml_find_all(nd, "./c:ser/c:explosion", ns = c(c = .nsmap[["c"]]))
      if (length(exp_nds) > 0) XL$DOUGHNUT_EXPLODED else XL$DOUGHNUT
    },
    "LinePlot" = {
      no_markers_nds <- xml2::xml_find_all(nd, 'c:ser/c:marker/c:symbol[@val="none"]',
                                            ns = c(c = .nsmap[["c"]]))
      has_markers <- length(no_markers_nds) == 0
      gv <- grouping_val()
      if (has_markers) {
        switch(gv,
          "standard"        = XL$LINE_MARKERS,
          "stacked"         = XL$LINE_MARKERS_STACKED,
          "percentStacked"  = XL$LINE_MARKERS_STACKED_100,
          XL$LINE_MARKERS
        )
      } else {
        switch(gv,
          "standard"        = XL$LINE,
          "stacked"         = XL$LINE_STACKED,
          "percentStacked"  = XL$LINE_STACKED_100,
          XL$LINE
        )
      }
    },
    "PiePlot" = {
      exp_nds <- xml2::xml_find_all(nd, "./c:ser/c:explosion", ns = c(c = .nsmap[["c"]]))
      if (length(exp_nds) > 0) XL$PIE_EXPLODED else XL$PIE
    },
    "RadarPlot" = {
      rs_nd <- xml2::xml_find_first(nd, "c:radarStyle", ns = c(c = .nsmap[["c"]]))
      rs <- if (inherits(rs_nd, "xml_missing")) NULL else xml2::xml_attr(rs_nd, "val")
      no_markers <- {
        ms_nds <- xml2::xml_find_all(nd, "c:ser/c:marker/c:symbol", ns = c(c = .nsmap[["c"]]))
        length(ms_nds) > 0 && xml2::xml_attr(ms_nds[[1]], "val") == "none"
      }
      if (!is.null(rs) && rs == "filled") return(XL$RADAR_FILLED)
      if (no_markers) XL$RADAR else XL$RADAR_MARKERS
    },
    "XyPlot" = {
      ss_nd <- xml2::xml_find_first(nd, "c:scatterStyle", ns = c(c = .nsmap[["c"]]))
      ss <- if (inherits(ss_nd, "xml_missing")) NULL else xml2::xml_attr(ss_nd, "val")
      no_line    <- length(xml2::xml_find_all(nd, "c:ser/c:spPr/a:ln/a:noFill",
                                              ns = c(c = .nsmap[["c"]], a = .nsmap[["a"]]))) > 0
      no_markers <- {
        ms <- xml2::xml_find_all(nd, "c:ser/c:marker/c:symbol", ns = c(c = .nsmap[["c"]]))
        length(ms) > 0 && xml2::xml_attr(ms[[1]], "val") == "none"
      }
      if (!is.null(ss) && ss == "lineMarker") {
        if (no_line) return(XL$XY_SCATTER)
        if (no_markers) return(XL$XY_SCATTER_LINES_NO_MARKERS)
        return(XL$XY_SCATTER_LINES)
      }
      if (!is.null(ss) && ss == "smoothMarker") {
        if (no_markers) return(XL$XY_SCATTER_SMOOTH_NO_MARKERS)
        return(XL$XY_SCATTER_SMOOTH)
      }
      XL$XY_SCATTER
    },
    stop(paste("chart_type not implemented for plot class:", class(plot)[1]), call. = FALSE)
  )
}


# ============================================================================
# Axis classes
# ============================================================================

#' Base class for chart axis domain objects
#' @noRd
BaseAxis <- R6::R6Class(
  "BaseAxis",
  public = list(
    initialize = function(xAx) {
      private$.element <- xAx
      private$.xAx     <- xAx
    }
  ),

  active = list(
    axis_title = function(value) {
      if (!missing(value)) return(invisible(NULL))
      AxisTitle$new(private$.element$get_or_add_title())
    },

    format = function(value) {
      if (!missing(value)) return(invisible(NULL))
      ChartFormat$new(private$.element)
    },

    has_major_gridlines = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) {
          private$.element$get_or_add_majorGridlines()
        } else {
          private$.element$.remove_majorGridlines()
        }
        return(invisible(value))
      }
      !is.null(private$.element$majorGridlines)
    },

    has_minor_gridlines = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) {
          private$.element$get_or_add_minorGridlines()
        } else {
          private$.element$.remove_minorGridlines()
        }
        return(invisible(value))
      }
      !is.null(private$.element$minorGridlines)
    },

    has_title = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) {
          private$.element$get_or_add_title()
        } else {
          private$.element$.remove_title()
        }
        return(invisible(value))
      }
      !is.null(private$.element$title)
    },

    major_gridlines = function(value) {
      if (!missing(value)) return(invisible(NULL))
      MajorGridlines$new(private$.element)
    },

    major_tick_mark = function(value) {
      tm <- private$.element$majorTickMark
      if (!missing(value)) {
        private$.element$.remove_majorTickMark()
        if (value != XL_TICK_MARK$CROSS) private$.element$.add_majorTickMark(val = value)
        return(invisible(value))
      }
      if (is.null(tm)) return(XL_TICK_MARK$CROSS)
      tm$val
    },

    maximum_scale = function(value) {
      if (!missing(value)) {
        sc <- private$.element$get_or_add_scaling()
        sc$maximum <- value
        return(invisible(value))
      }
      sc <- private$.element$scaling
      if (is.null(sc)) return(NULL)
      sc$maximum
    },

    minimum_scale = function(value) {
      if (!missing(value)) {
        sc <- private$.element$get_or_add_scaling()
        sc$minimum <- value
        return(invisible(value))
      }
      sc <- private$.element$scaling
      if (is.null(sc)) return(NULL)
      sc$minimum
    },

    minor_tick_mark = function(value) {
      tm <- private$.element$minorTickMark
      if (!missing(value)) {
        private$.element$.remove_minorTickMark()
        if (value != XL_TICK_MARK$CROSS) private$.element$.add_minorTickMark(val = value)
        return(invisible(value))
      }
      if (is.null(tm)) return(XL_TICK_MARK$CROSS)
      tm$val
    },

    reverse_order = function(value) {
      if (!missing(value)) {
        private$.element$orientation <- if (isTRUE(value)) "maxMin" else "minMax"
        return(invisible(value))
      }
      private$.element$orientation == "maxMin"
    },

    tick_labels = function(value) {
      if (!missing(value)) return(invisible(NULL))
      TickLabels$new(private$.element)
    },

    tick_label_position = function(value) {
      tlp <- private$.element$tickLblPos
      if (!missing(value)) {
        tlp_elm <- private$.element$get_or_add_tickLblPos()
        tlp_elm$val <- value
        return(invisible(value))
      }
      if (is.null(tlp)) return(XL_TICK_LABEL_POSITION$NEXT_TO_AXIS)
      v <- tlp$val
      if (is.null(v)) return(XL_TICK_LABEL_POSITION$NEXT_TO_AXIS)
      v
    },

    visible = function(value) {
      del <- private$.element$delete_
      if (!missing(value)) {
        if (!is.logical(value)) stop("visible must be TRUE or FALSE", call. = FALSE)
        del_elm <- private$.element$get_or_add_delete_()
        del_elm$val <- !value
        return(invisible(value))
      }
      if (is.null(del)) return(TRUE)
      !isTRUE(del$val)
    }
  ),

  private = list(.element = NULL, .xAx = NULL)
)

#' Category axis domain object
#' @noRd
CategoryAxis <- R6::R6Class(
  "CategoryAxis",
  inherit = BaseAxis,
  active = list(
    category_type = function() XL_CATEGORY_TYPE$CATEGORY_SCALE
  )
)

#' Date axis domain object
#' @noRd
DateAxis <- R6::R6Class(
  "DateAxis",
  inherit = BaseAxis,
  active = list(
    category_type = function() XL_CATEGORY_TYPE$TIME_SCALE
  )
)

#' Value axis domain object
#' @noRd
ValueAxis <- R6::R6Class(
  "ValueAxis",
  inherit = BaseAxis,
  active = list(
    crosses = function(value) {
      cr_xAx <- private$.cross_xAx()
      if (!missing(value)) {
        cr_xAx$.remove_crosses()
        cr_xAx$.remove_crossesAt()
        if (value == XL_AXIS_CROSSES$CUSTOM) {
          cr_xAx$.add_crossesAt(val = 0.0)
        } else {
          cr_xAx$.add_crosses(val = value)
        }
        return(invisible(value))
      }
      crosses_elm <- cr_xAx$crosses
      if (is.null(crosses_elm)) return(XL_AXIS_CROSSES$CUSTOM)
      crosses_elm$val
    },

    crosses_at = function(value) {
      cr_xAx <- private$.cross_xAx()
      if (!missing(value)) {
        cr_xAx$.remove_crosses()
        cr_xAx$.remove_crossesAt()
        if (!is.null(value)) cr_xAx$.add_crossesAt(val = value)
        return(invisible(value))
      }
      cat_nd <- cr_xAx$crossesAt
      if (is.null(cat_nd)) return(NULL)
      cat_nd$val
    },

    major_unit = function(value) {
      if (!missing(value)) {
        if (is.null(value)) {
          private$.element$.remove_majorUnit()
        } else {
          private$.element$.add_majorUnit(val = value)
        }
        return(invisible(value))
      }
      mu <- private$.element$majorUnit
      if (is.null(mu)) return(NULL)
      mu$val
    },

    minor_unit = function(value) {
      if (!missing(value)) {
        if (is.null(value)) {
          private$.element$.remove_minorUnit()
        } else {
          private$.element$.add_minorUnit(val = value)
        }
        return(invisible(value))
      }
      mu <- private$.element$minorUnit
      if (is.null(mu)) return(NULL)
      mu$val
    }
  ),

  private = list(
    # Return the axis element that crosses this value axis
    .cross_xAx = function() {
      crossAx_elm <- private$.element$crossAx
      if (is.null(crossAx_elm)) return(private$.element)
      crossAx_id <- crossAx_elm$val
      # Find sibling axis with matching axId
      parent_nd <- xml2::xml_parent(private$.element$get_node())
      ax_nds <- xml2::xml_find_all(parent_nd,
        "(c:catAx|c:valAx|c:dateAx)/c:axId",
        ns = c(c = .nsmap[["c"]])
      )
      for (ax_id_nd in ax_nds) {
        if (as.integer(xml2::xml_attr(ax_id_nd, "val")) == crossAx_id) {
          return(wrap_element(xml2::xml_parent(ax_id_nd)))
        }
      }
      private$.element
    }
  )
)


# ============================================================================
# AxisTitle, ChartTitle
# ============================================================================

#' Axis title domain object
#' @noRd
AxisTitle <- R6::R6Class(
  "AxisTitle",
  public = list(
    initialize = function(title) {
      private$.element <- title
      private$.title   <- title
    }
  ),
  active = list(
    format = function(value) {
      if (!missing(value)) return(invisible(NULL))
      ChartFormat$new(private$.element)
    },
    has_text_frame = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) {
          private$.title$get_or_add_tx_rich()
        } else {
          private$.title$.remove_tx()
        }
        return(invisible(value))
      }
      !is.null(private$.title$tx_rich)
    },
    text_frame = function(value) {
      if (!missing(value)) return(invisible(NULL))
      rich <- private$.title$get_or_add_tx_rich()
      # get_or_add_tx_rich returns tx; we need rich inside it
      rich_nd <- xml2::xml_find_first(rich$get_node(), "c:rich", ns = c(c = .nsmap[["c"]]))
      if (inherits(rich_nd, "xml_missing")) {
        rich_nd <- xml2::xml_find_first(rich$get_node(), "c:rich", ns = c(c = .nsmap[["c"]]))
      }
      TextFrame$new(wrap_element(rich_nd), self)
    }
  ),
  private = list(.element = NULL, .title = NULL)
)

#' Chart title domain object
#' @noRd
ChartTitle <- R6::R6Class(
  "ChartTitle",
  public = list(
    initialize = function(title) {
      private$.element <- title
      private$.title   <- title
    }
  ),
  active = list(
    format = function(value) {
      if (!missing(value)) return(invisible(NULL))
      ChartFormat$new(private$.element)
    },
    has_text_frame = function(value) {
      if (!missing(value)) {
        if (isTRUE(value)) {
          private$.title$get_or_add_tx_rich()
        } else {
          private$.title$.remove_tx()
        }
        return(invisible(value))
      }
      !is.null(private$.title$tx_rich)
    },
    text_frame = function(value) {
      if (!missing(value)) return(invisible(NULL))
      tx <- private$.title$get_or_add_tx_rich()
      rich_nd <- xml2::xml_find_first(tx$get_node(), "c:rich", ns = c(c = .nsmap[["c"]]))
      TextFrame$new(wrap_element(rich_nd), self)
    }
  ),
  private = list(.element = NULL, .title = NULL)
)


# ============================================================================
# MajorGridlines
# ============================================================================

#' Major gridlines domain object
#' @noRd
MajorGridlines <- R6::R6Class(
  "MajorGridlines",
  public = list(
    initialize = function(xAx) {
      private$.xAx <- xAx
    }
  ),
  active = list(
    format = function(value) {
      if (!missing(value)) return(invisible(NULL))
      gridlines <- private$.xAx$get_or_add_majorGridlines()
      ChartFormat$new(gridlines)
    }
  ),
  private = list(.xAx = NULL)
)


# ============================================================================
# TickLabels
# ============================================================================

#' Tick label formatting object for a chart axis
#' @noRd
TickLabels <- R6::R6Class(
  "TickLabels",
  public = list(
    initialize = function(xAx_elm) {
      private$.element <- xAx_elm
    }
  ),

  active = list(
    font = function(value) {
      if (!missing(value)) return(invisible(NULL))
      defRPr <- private$.element$defRPr
      Font$new(defRPr)
    },

    number_format = function(value) {
      if (!missing(value)) {
        { nm <- private$.element$get_or_add_numFmt(); nm$formatCode <- value }
        self$number_format_is_linked <- FALSE
        return(invisible(value))
      }
      numFmt <- private$.element$numFmt
      if (is.null(numFmt)) return("General")
      numFmt$formatCode
    },

    number_format_is_linked = function(value) {
      if (!missing(value)) {
        { nm <- private$.element$get_or_add_numFmt(); nm$sourceLinked <- value }
        return(invisible(value))
      }
      numFmt <- private$.element$numFmt
      if (is.null(numFmt)) return(FALSE)
      sl <- numFmt$sourceLinked
      if (is.null(sl)) return(TRUE)
      sl
    },

    offset = function(value) {
      lo <- private$.element$lblOffset
      if (!missing(value)) {
        nd <- private$.element$get_node()
        tag <- xml2::xml_name(nd, ns = xml2::xml_ns(nd))
        if (tag != "c:catAx") stop("only a category axis has an offset", call. = FALSE)
        private$.element$.remove_lblOffset()
        if (value != 100L) {
          { lo_elm <- private$.element$.add_lblOffset(); lo_elm$val <- value }
        }
        return(invisible(value))
      }
      if (is.null(lo)) return(100L)
      lo$val
    }
  ),

  private = list(.element = NULL)
)


# ============================================================================
# Legend
# ============================================================================

#' Chart legend domain object
#' @noRd
Legend <- R6::R6Class(
  "Legend",
  public = list(
    initialize = function(legend_elm) {
      private$.element <- legend_elm
    }
  ),

  active = list(
    font = function(value) {
      if (!missing(value)) return(invisible(NULL))
      defRPr <- private$.element$defRPr
      Font$new(defRPr)
    },

    horz_offset = function(value) {
      if (!missing(value)) {
        private$.element$horz_offset <- value
        return(invisible(value))
      }
      private$.element$horz_offset
    },

    include_in_layout = function(value) {
      ovl <- private$.element$overlay
      if (!missing(value)) {
        if (is.null(value)) {
          private$.element$.remove_overlay()
        } else {
          ovl_elm <- private$.element$get_or_add_overlay(); ovl_elm$val <- isTRUE(value)
        }
        return(invisible(value))
      }
      if (is.null(ovl)) return(TRUE)
      ovl$val
    },

    position = function(value) {
      lp <- private$.element$legendPos
      if (!missing(value)) {
        lp_elm <- private$.element$get_or_add_legendPos()
        lp_elm$val <- value
        return(invisible(value))
      }
      if (is.null(lp)) return(XL_LEGEND_POSITION$RIGHT)
      lp$val
    }
  ),

  private = list(.element = NULL)
)


# ============================================================================
# Chart — the main user-facing chart domain object
# ============================================================================

#' Chart domain object
#'
#' Provides access to chart properties, title, legend, plots, series, and axes.
#' Access via `graphic_frame$chart`.
#'
#' @noRd
Chart <- R6::R6Class(
  "Chart",
  public = list(
    initialize = function(chartSpace, chart_part) {
      private$.element     <- chartSpace
      private$.chartSpace  <- chartSpace
      private$.chart_part  <- chart_part
    }
  ),

  active = list(
    # The ChartPart that owns this chart
    part = function() private$.chart_part,

    category_axis = function(value) {
      if (!missing(value)) return(invisible(NULL))
      cat_lst <- private$.chartSpace$catAx_lst
      if (length(cat_lst) > 0) return(CategoryAxis$new(cat_lst[[1]]))
      date_lst <- private$.chartSpace$dateAx_lst
      if (length(date_lst) > 0) return(DateAxis$new(date_lst[[1]]))
      val_lst <- private$.chartSpace$valAx_lst
      if (length(val_lst) > 0) return(ValueAxis$new(val_lst[[1]]))
      stop("chart has no category axis", call. = FALSE)
    },

    chart_style = function(value) {
      if (!missing(value)) {
        private$.chartSpace$.remove_style()
        if (!is.null(value)) private$.chartSpace$.add_style(val = value)
        return(invisible(value))
      }
      sty <- private$.chartSpace$style
      if (is.null(sty)) return(NULL)
      sty$val
    },

    chart_title = function(value) {
      if (!missing(value)) return(invisible(NULL))
      ChartTitle$new(private$.chartSpace$get_or_add_title())
    },

    chart_type = function() {
      first_plot <- self$plots[[1]]
      plot_type_inspector(first_plot)
    },

    has_legend = function(value) {
      if (!missing(value)) {
        ch_elm <- private$.chartSpace$chart
        ch_elm$has_legend <- isTRUE(value)
        return(invisible(value))
      }
      private$.chartSpace$chart$has_legend
    },

    has_title = function(value) {
      if (!missing(value)) {
        chart_elm <- private$.chartSpace$chart
        if (!isTRUE(value)) {
          chart_elm$.remove_title()
          atd_elm <- chart_elm$get_or_add_autoTitleDeleted()
          atd_elm$val <- TRUE
        } else {
          chart_elm$get_or_add_title()
        }
        return(invisible(value))
      }
      !is.null(private$.chartSpace$chart$title)
    },

    legend = function(value) {
      if (!missing(value)) return(invisible(NULL))
      legend_elm <- private$.chartSpace$chart$legend
      if (is.null(legend_elm)) return(NULL)
      Legend$new(legend_elm)
    },

    plots = function(value) {
      if (!missing(value)) return(invisible(NULL))
      xCharts <- private$.chartSpace$plotArea$xCharts
      lapply(xCharts, function(xc) plot_factory(xc, self))
    },

    series = function(value) {
      if (!missing(value)) return(invisible(NULL))
      SeriesCollection$new(private$.chartSpace$plotArea)
    },

    value_axis = function(value) {
      if (!missing(value)) return(invisible(NULL))
      val_lst <- private$.chartSpace$valAx_lst
      if (length(val_lst) == 0) stop("chart has no value axis", call. = FALSE)
      idx <- if (length(val_lst) > 1) 2L else 1L
      ValueAxis$new(val_lst[[idx]])
    }
  ),

  private = list(
    .element    = NULL,
    .chartSpace = NULL,
    .chart_part = NULL
  )
)
