# Chart XML generation for various chart types.
#
# Ported from python-pptx/src/pptx/chart/xmlwriter.py.
# Generates the c:chartSpace XML string for new charts.


# Escape XML special characters in a string.
.xml_escape <- function(s) {
  s <- gsub("&", "&amp;", as.character(s), fixed = TRUE)
  s <- gsub("<", "&lt;",  s, fixed = TRUE)
  s <- gsub(">", "&gt;",  s, fixed = TRUE)
  s
}


# ============================================================================
# Factory function
# ============================================================================

#' Return the appropriate chart XML writer for chart_type
#'
#' @param chart_type An integer from XL_CHART_TYPE, e.g. `XL_CHART_TYPE$COLUMN_CLUSTERED`.
#' @param chart_data A `CategoryChartData`, `XyChartData`, or `BubbleChartData` object.
#' @return An R6 writer object with an `$xml` active binding returning the XML string.
#' @noRd
chart_xml_writer <- function(chart_type, chart_data) {
  XL <- XL_CHART_TYPE
  dispatch <- list(
    `area`          = list(XL$AREA, XL$AREA_STACKED, XL$AREA_STACKED_100),
    `bar`           = list(XL$BAR_CLUSTERED, XL$BAR_STACKED, XL$BAR_STACKED_100,
                           XL$COLUMN_CLUSTERED, XL$COLUMN_STACKED, XL$COLUMN_STACKED_100),
    `bubble`        = list(XL$BUBBLE, XL$BUBBLE_THREE_D_EFFECT),
    `doughnut`      = list(XL$DOUGHNUT, XL$DOUGHNUT_EXPLODED),
    `line`          = list(XL$LINE, XL$LINE_MARKERS, XL$LINE_MARKERS_STACKED,
                           XL$LINE_MARKERS_STACKED_100, XL$LINE_STACKED, XL$LINE_STACKED_100),
    `pie`           = list(XL$PIE, XL$PIE_EXPLODED),
    `radar`         = list(XL$RADAR, XL$RADAR_FILLED, XL$RADAR_MARKERS),
    `xy`            = list(XL$XY_SCATTER, XL$XY_SCATTER_LINES, XL$XY_SCATTER_LINES_NO_MARKERS,
                           XL$XY_SCATTER_SMOOTH, XL$XY_SCATTER_SMOOTH_NO_MARKERS)
  )
  writer_type <- NULL
  for (nm in names(dispatch)) {
    if (chart_type %in% unlist(dispatch[[nm]])) {
      writer_type <- nm
      break
    }
  }
  if (is.null(writer_type))
    stop(sprintf("XML writer for chart type %s not yet implemented", chart_type), call. = FALSE)

  switch(writer_type,
    area     = .AreaChartXmlWriter$new(chart_type, chart_data),
    bar      = .BarChartXmlWriter$new(chart_type, chart_data),
    bubble   = .BubbleChartXmlWriter$new(chart_type, chart_data),
    doughnut = .DoughnutChartXmlWriter$new(chart_type, chart_data),
    line     = .LineChartXmlWriter$new(chart_type, chart_data),
    pie      = .PieChartXmlWriter$new(chart_type, chart_data),
    radar    = .RadarChartXmlWriter$new(chart_type, chart_data),
    xy       = .XyChartXmlWriter$new(chart_type, chart_data)
  )
}


# ============================================================================
# Base chart XML writer
# ============================================================================

.BaseChartXmlWriter <- R6::R6Class(
  ".BaseChartXmlWriter",

  public = list(
    initialize = function(chart_type, chart_data) {
      private$.chart_type <- chart_type
      private$.chart_data <- chart_data
    }
  ),

  active = list(
    xml = function() stop("must be implemented by subclass", call. = FALSE)
  ),

  private = list(
    .chart_type = NULL,
    .chart_data = NULL
  )
)


# ============================================================================
# Series XML writers (helpers, not exported)
# ============================================================================

# ---------- _CategorySeriesXmlWriter -----------------------------------------

.CategorySeriesXmlWriter <- R6::R6Class(
  ".CategorySeriesXmlWriter",

  public = list(
    initialize = function(series, date_1904 = FALSE) {
      private$.series    <- series
      private$.date_1904 <- date_1904
    },

    # <c:tx> element XML for the series name.
    tx_xml = function() {
      paste0(
        "          <c:tx>\n",
        "            <c:strRef>\n",
        "              <c:f>", private$.series$name_ref, "</c:f>\n",
        "              <c:strCache>\n",
        '                <c:ptCount val="1"/>\n',
        '                <c:pt idx="0">\n',
        "                  <c:v>", .xml_escape(private$.series$name), "</c:v>\n",
        "                </c:pt>\n",
        "              </c:strCache>\n",
        "            </c:strRef>\n",
        "          </c:tx>\n"
      )
    },

    # <c:cat> element XML.
    cat_xml = function() {
      cats <- private$.series$categories
      if (cats$are_numeric) {
        return(private$.numRef_cat_xml())
      }
      if (cats$depth == 1L) {
        return(private$.strRef_cat_xml())
      }
      private$.multiLvl_cat_xml()
    },

    # <c:val> element XML.
    val_xml = function() {
      private$.val_xml_impl()
    }
  ),

  private = list(
    .series    = NULL,
    .date_1904 = FALSE,

    .numRef_cat_xml = function() {
      cats     <- private$.series$categories
      cat_pt   <- private$.cat_num_pt_xml()
      paste0(
        "          <c:cat>\n",
        "            <c:numRef>\n",
        "              <c:f>", private$.series$categories_ref, "</c:f>\n",
        "              <c:numCache>\n",
        "                <c:formatCode>", cats$number_format, "</c:formatCode>\n",
        sprintf('                <c:ptCount val="%d"/>\n', cats$leaf_count),
        cat_pt,
        "              </c:numCache>\n",
        "            </c:numRef>\n",
        "          </c:cat>\n"
      )
    },

    .strRef_cat_xml = function() {
      cats <- private$.series$categories
      paste0(
        "          <c:cat>\n",
        "            <c:strRef>\n",
        "              <c:f>", private$.series$categories_ref, "</c:f>\n",
        "              <c:strCache>\n",
        sprintf('                <c:ptCount val="%d"/>\n', cats$leaf_count),
        private$.cat_pt_xml(),
        "              </c:strCache>\n",
        "            </c:strRef>\n",
        "          </c:cat>\n"
      )
    },

    .multiLvl_cat_xml = function() {
      cats <- private$.series$categories
      paste0(
        "          <c:cat>\n",
        "            <c:multiLvlStrRef>\n",
        "              <c:f>", private$.series$categories_ref, "</c:f>\n",
        "              <c:multiLvlStrCache>\n",
        sprintf('                <c:ptCount val="%d"/>\n', cats$leaf_count),
        private$.lvl_xml(cats),
        "              </c:multiLvlStrCache>\n",
        "            </c:multiLvlStrRef>\n",
        "          </c:cat>\n"
      )
    },

    .cat_pt_xml = function() {
      cats <- private$.series$categories
      n    <- length(cats)
      xml  <- ""
      for (i in seq_len(n)) {
        cat   <- cats$get(i)
        idx   <- i - 1L
        label <- .xml_escape(as.character(cat$label))
        xml   <- paste0(xml, sprintf(
          '                <c:pt idx="%d">\n                  <c:v>%s</c:v>\n                </c:pt>\n',
          idx, label
        ))
      }
      xml
    },

    .cat_num_pt_xml = function() {
      cats <- private$.series$categories
      n    <- length(cats)
      xml  <- ""
      for (i in seq_len(n)) {
        cat     <- cats$get(i)
        idx     <- i - 1L
        lbl_str <- cat$numeric_str_val(date_1904 = private$.date_1904)
        xml <- paste0(xml, sprintf(
          '                <c:pt idx="%d">\n                  <c:v>%s</c:v>\n                </c:pt>\n',
          idx, lbl_str
        ))
      }
      xml
    },

    .lvl_xml = function(cats) {
      xml    <- ""
      levels <- cats$levels
      for (level in levels) {
        lvl_pt <- ""
        for (entry in level) {
          lvl_pt <- paste0(lvl_pt, sprintf(
            '                  <c:pt idx="%d">\n                    <c:v>%s</c:v>\n                  </c:pt>\n',
            entry$off, .xml_escape(as.character(entry$name))
          ))
        }
        xml <- paste0(xml, "                <c:lvl>\n", lvl_pt, "                </c:lvl>\n")
      }
      xml
    },

    .val_pt_xml = function() {
      values <- private$.series$values
      xml    <- ""
      for (i in seq_along(values)) {
        v <- values[i]
        if (is.na(v)) next
        xml <- paste0(xml, sprintf(
          '                <c:pt idx="%d">\n                  <c:v>%s</c:v>\n                </c:pt>\n',
          i - 1L, v
        ))
      }
      xml
    },

    .val_xml_impl = function() {
      s <- private$.series
      paste0(
        "          <c:val>\n",
        "            <c:numRef>\n",
        "              <c:f>", s$values_ref, "</c:f>\n",
        "              <c:numCache>\n",
        "                <c:formatCode>", s$number_format, "</c:formatCode>\n",
        sprintf('                <c:ptCount val="%d"/>\n', s$n_points),
        private$.val_pt_xml(),
        "              </c:numCache>\n",
        "            </c:numRef>\n",
        "          </c:val>\n"
      )
    }
  )
)

# ---------- _XySeriesXmlWriter -----------------------------------------------

.pt_xml <- function(values) {
  xml <- sprintf('                <c:ptCount val="%d"/>\n', length(values))
  for (i in seq_along(values)) {
    v <- values[i]
    if (is.na(v)) next
    xml <- paste0(xml, sprintf(
      '                <c:pt idx="%d">\n                  <c:v>%s</c:v>\n                </c:pt>\n',
      i - 1L, v
    ))
  }
  xml
}

.numRef_xml <- function(wksht_ref, number_format, values) {
  paste0(
    "            <c:numRef>\n",
    "              <c:f>", wksht_ref, "</c:f>\n",
    "              <c:numCache>\n",
    "                <c:formatCode>", number_format, "</c:formatCode>\n",
    .pt_xml(values),
    "              </c:numCache>\n",
    "            </c:numRef>\n"
  )
}

.XySeriesXmlWriter <- R6::R6Class(
  ".XySeriesXmlWriter",

  public = list(
    initialize = function(series, date_1904 = FALSE) {
      private$.series    <- series
      private$.date_1904 <- date_1904
    },

    tx_xml = function() {
      paste0(
        "          <c:tx>\n",
        "            <c:strRef>\n",
        "              <c:f>", private$.series$name_ref, "</c:f>\n",
        "              <c:strCache>\n",
        '                <c:ptCount val="1"/>\n',
        '                <c:pt idx="0">\n',
        "                  <c:v>", .xml_escape(private$.series$name), "</c:v>\n",
        "                </c:pt>\n",
        "              </c:strCache>\n",
        "            </c:strRef>\n",
        "          </c:tx>\n"
      )
    },

    xVal_xml = function() {
      s <- private$.series
      paste0(
        "          <c:xVal>\n",
        .numRef_xml(s$x_values_ref, s$number_format, s$x_values),
        "          </c:xVal>\n"
      )
    },

    yVal_xml = function() {
      s <- private$.series
      paste0(
        "          <c:yVal>\n",
        .numRef_xml(s$y_values_ref, s$number_format, s$y_values),
        "          </c:yVal>\n"
      )
    }
  ),

  private = list(
    .series    = NULL,
    .date_1904 = FALSE
  )
)

# ---------- _BubbleSeriesXmlWriter -------------------------------------------

.BubbleSeriesXmlWriter <- R6::R6Class(
  ".BubbleSeriesXmlWriter",
  inherit = .XySeriesXmlWriter,

  public = list(
    bubbleSize_xml = function() {
      s <- private$.series
      paste0(
        "          <c:bubbleSize>\n",
        .numRef_xml(s$bubble_sizes_ref, s$number_format, s$bubble_sizes),
        "          </c:bubbleSize>\n"
      )
    }
  )
)


# ============================================================================
# _AreaChartXmlWriter
# ============================================================================

.AreaChartXmlWriter <- R6::R6Class(
  ".AreaChartXmlWriter",
  inherit = .BaseChartXmlWriter,

  active = list(
    xml = function() {
      paste0(
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n",
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
        '  <c:date1904 val="0"/>\n',
        '  <c:roundedCorners val="0"/>\n',
        "  <c:chart>\n",
        '    <c:autoTitleDeleted val="0"/>\n',
        "    <c:plotArea>\n",
        "      <c:layout/>\n",
        "      <c:areaChart>\n",
        private$.grouping_xml(),
        '        <c:varyColors val="0"/>\n',
        private$.ser_xml(),
        "        <c:dLbls>\n",
        '          <c:showLegendKey val="0"/>\n',
        '          <c:showVal val="0"/>\n',
        '          <c:showCatName val="0"/>\n',
        '          <c:showSerName val="0"/>\n',
        '          <c:showPercent val="0"/>\n',
        '          <c:showBubbleSize val="0"/>\n',
        "        </c:dLbls>\n",
        '        <c:axId val="-2101159928"/>\n',
        '        <c:axId val="-2100718248"/>\n',
        "      </c:areaChart>\n",
        private$.cat_ax_xml(),
        "      <c:valAx>\n",
        '        <c:axId val="-2100718248"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="l"/>\n',
        "        <c:majorGridlines/>\n",
        '        <c:numFmt formatCode="General" sourceLinked="1"/>\n',
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="-2101159928"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:crossBetween val="midCat"/>\n',
        "      </c:valAx>\n",
        "    </c:plotArea>\n",
        "    <c:legend>\n",
        '      <c:legendPos val="r"/>\n',
        "      <c:layout/>\n",
        '      <c:overlay val="0"/>\n',
        "    </c:legend>\n",
        '    <c:plotVisOnly val="1"/>\n',
        '    <c:dispBlanksAs val="zero"/>\n',
        '    <c:showDLblsOverMax val="0"/>\n',
        "  </c:chart>\n",
        "  <c:txPr>\n",
        "    <a:bodyPr/>\n",
        "    <a:lstStyle/>\n",
        "    <a:p>\n",
        "      <a:pPr>\n",
        '        <a:defRPr sz="1800"/>\n',
        "      </a:pPr>\n",
        "      <a:endParaRPr/>\n",
        "    </a:p>\n",
        "  </c:txPr>\n",
        "</c:chartSpace>\n"
      )
    }
  ),

  private = list(
    .grouping_xml = function() {
      XL  <- XL_CHART_TYPE
      val <- switch(
        as.character(private$.chart_type),
        `1`   = "standard",   # AREA
        `76`  = "stacked",    # AREA_STACKED
        `77`  = "percentStacked", # AREA_STACKED_100
        stop("unexpected area chart type", call. = FALSE)
      )
      sprintf('        <c:grouping val="%s"/>\n', val)
    },

    .cat_ax_xml = function() {
      cats <- private$.chart_data$categories
      if (cats$are_dates) {
        return(paste0(
          "      <c:dateAx>\n",
          '        <c:axId val="-2101159928"/>\n',
          "        <c:scaling>\n",
          '          <c:orientation val="minMax"/>\n',
          "        </c:scaling>\n",
          '        <c:delete val="0"/>\n',
          '        <c:axPos val="b"/>\n',
          sprintf('        <c:numFmt formatCode="%s" sourceLinked="1"/>\n', cats$number_format),
          '        <c:majorTickMark val="out"/>\n',
          '        <c:minorTickMark val="none"/>\n',
          '        <c:tickLblPos val="nextTo"/>\n',
          '        <c:crossAx val="-2100718248"/>\n',
          '        <c:crosses val="autoZero"/>\n',
          '        <c:auto val="1"/>\n',
          '        <c:lblOffset val="100"/>\n',
          '        <c:baseTimeUnit val="days"/>\n',
          "      </c:dateAx>\n"
        ))
      }
      paste0(
        "      <c:catAx>\n",
        '        <c:axId val="-2101159928"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="b"/>\n',
        '        <c:numFmt formatCode="General" sourceLinked="1"/>\n',
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="-2100718248"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:auto val="1"/>\n',
        '        <c:lblAlgn val="ctr"/>\n',
        '        <c:lblOffset val="100"/>\n',
        '        <c:noMultiLvlLbl val="0"/>\n',
        "      </c:catAx>\n"
      )
    },

    .ser_xml = function() {
      xml <- ""
      for (series in private$.chart_data$to_list()) {
        w   <- .CategorySeriesXmlWriter$new(series)
        xml <- paste0(xml,
          "        <c:ser>\n",
          sprintf('          <c:idx val="%d"/>\n', series$index),
          sprintf('          <c:order val="%d"/>\n', series$index),
          w$tx_xml(),
          w$cat_xml(),
          w$val_xml(),
          "        </c:ser>\n"
        )
      }
      xml
    }
  )
)


# ============================================================================
# _BarChartXmlWriter (bar + column)
# ============================================================================

.BarChartXmlWriter <- R6::R6Class(
  ".BarChartXmlWriter",
  inherit = .BaseChartXmlWriter,

  active = list(
    xml = function() {
      paste0(
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n",
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
        '  <c:date1904 val="0"/>\n',
        "  <c:chart>\n",
        '    <c:autoTitleDeleted val="0"/>\n',
        "    <c:plotArea>\n",
        "      <c:barChart>\n",
        private$.barDir_xml(),
        private$.grouping_xml(),
        private$.ser_xml(),
        private$.overlap_xml(),
        '        <c:axId val="-2068027336"/>\n',
        '        <c:axId val="-2113994440"/>\n',
        "      </c:barChart>\n",
        private$.cat_ax_xml(),
        "      <c:valAx>\n",
        '        <c:axId val="-2113994440"/>\n',
        "        <c:scaling/>\n",
        '        <c:delete val="0"/>\n',
        sprintf('        <c:axPos val="%s"/>\n', private$.val_ax_pos()),
        "        <c:majorGridlines/>\n",
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="-2068027336"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        "      </c:valAx>\n",
        "    </c:plotArea>\n",
        '    <c:dispBlanksAs val="gap"/>\n',
        "  </c:chart>\n",
        "  <c:txPr>\n",
        "    <a:bodyPr/>\n",
        "    <a:lstStyle/>\n",
        "    <a:p>\n",
        "      <a:pPr>\n",
        '        <a:defRPr sz="1800"/>\n',
        "      </a:pPr>\n",
        '      <a:endParaRPr lang="en-US"/>\n',
        "    </a:p>\n",
        "  </c:txPr>\n",
        "</c:chartSpace>\n"
      )
    }
  ),

  private = list(
    .barDir_xml = function() {
      XL <- XL_CHART_TYPE
      bar_types <- c(XL$BAR_CLUSTERED, XL$BAR_STACKED, XL$BAR_STACKED_100)
      col_types <- c(XL$COLUMN_CLUSTERED, XL$COLUMN_STACKED, XL$COLUMN_STACKED_100)
      if (private$.chart_type %in% bar_types) return('        <c:barDir val="bar"/>\n')
      if (private$.chart_type %in% col_types) return('        <c:barDir val="col"/>\n')
      stop("unexpected bar/column chart type", call. = FALSE)
    },

    .grouping_xml = function() {
      XL <- XL_CHART_TYPE
      clustered <- c(XL$BAR_CLUSTERED, XL$COLUMN_CLUSTERED)
      stacked   <- c(XL$BAR_STACKED, XL$COLUMN_STACKED)
      pct       <- c(XL$BAR_STACKED_100, XL$COLUMN_STACKED_100)
      if (private$.chart_type %in% clustered) return('        <c:grouping val="clustered"/>\n')
      if (private$.chart_type %in% stacked)   return('        <c:grouping val="stacked"/>\n')
      if (private$.chart_type %in% pct)       return('        <c:grouping val="percentStacked"/>\n')
      stop("unexpected grouping for bar chart type", call. = FALSE)
    },

    .overlap_xml = function() {
      XL <- XL_CHART_TYPE
      stacked <- c(XL$BAR_STACKED, XL$BAR_STACKED_100, XL$COLUMN_STACKED, XL$COLUMN_STACKED_100)
      if (private$.chart_type %in% stacked) return('        <c:overlap val="100"/>\n')
      ""
    },

    .cat_ax_pos = function() {
      XL  <- XL_CHART_TYPE
      bar <- c(XL$BAR_CLUSTERED, XL$BAR_STACKED, XL$BAR_STACKED_100)
      col <- c(XL$COLUMN_CLUSTERED, XL$COLUMN_STACKED, XL$COLUMN_STACKED_100)
      if (private$.chart_type %in% bar) return("l")
      if (private$.chart_type %in% col) return("b")
      stop("unexpected chart type for cat_ax_pos", call. = FALSE)
    },

    .val_ax_pos = function() {
      XL  <- XL_CHART_TYPE
      bar <- c(XL$BAR_CLUSTERED, XL$BAR_STACKED, XL$BAR_STACKED_100)
      col <- c(XL$COLUMN_CLUSTERED, XL$COLUMN_STACKED, XL$COLUMN_STACKED_100)
      if (private$.chart_type %in% bar) return("b")
      if (private$.chart_type %in% col) return("l")
      stop("unexpected chart type for val_ax_pos", call. = FALSE)
    },

    .cat_ax_xml = function() {
      cats    <- private$.chart_data$categories
      ax_pos  <- private$.cat_ax_pos()
      if (cats$are_dates) {
        return(paste0(
          "      <c:dateAx>\n",
          '        <c:axId val="-2068027336"/>\n',
          "        <c:scaling>\n",
          '          <c:orientation val="minMax"/>\n',
          "        </c:scaling>\n",
          '        <c:delete val="0"/>\n',
          sprintf('        <c:axPos val="%s"/>\n', ax_pos),
          sprintf('        <c:numFmt formatCode="%s" sourceLinked="1"/>\n', cats$number_format),
          '        <c:majorTickMark val="out"/>\n',
          '        <c:minorTickMark val="none"/>\n',
          '        <c:tickLblPos val="nextTo"/>\n',
          '        <c:crossAx val="-2113994440"/>\n',
          '        <c:crosses val="autoZero"/>\n',
          '        <c:auto val="1"/>\n',
          '        <c:lblOffset val="100"/>\n',
          '        <c:baseTimeUnit val="days"/>\n',
          "      </c:dateAx>\n"
        ))
      }
      paste0(
        "      <c:catAx>\n",
        '        <c:axId val="-2068027336"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        sprintf('        <c:axPos val="%s"/>\n', ax_pos),
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="-2113994440"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:auto val="1"/>\n',
        '        <c:lblAlgn val="ctr"/>\n',
        '        <c:lblOffset val="100"/>\n',
        '        <c:noMultiLvlLbl val="0"/>\n',
        "      </c:catAx>\n"
      )
    },

    .ser_xml = function() {
      xml <- ""
      for (series in private$.chart_data$to_list()) {
        w   <- .CategorySeriesXmlWriter$new(series)
        xml <- paste0(xml,
          "        <c:ser>\n",
          sprintf('          <c:idx val="%d"/>\n', series$index),
          sprintf('          <c:order val="%d"/>\n', series$index),
          w$tx_xml(),
          w$cat_xml(),
          w$val_xml(),
          "        </c:ser>\n"
        )
      }
      xml
    }
  )
)


# ============================================================================
# _DoughnutChartXmlWriter
# ============================================================================

.DoughnutChartXmlWriter <- R6::R6Class(
  ".DoughnutChartXmlWriter",
  inherit = .BaseChartXmlWriter,

  active = list(
    xml = function() {
      paste0(
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n",
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
        '  <c:date1904 val="0"/>\n',
        '  <c:roundedCorners val="0"/>\n',
        "  <c:chart>\n",
        '    <c:autoTitleDeleted val="0"/>\n',
        "    <c:plotArea>\n",
        "      <c:layout/>\n",
        "      <c:doughnutChart>\n",
        '        <c:varyColors val="1"/>\n',
        private$.ser_xml(),
        "        <c:dLbls>\n",
        '          <c:showLegendKey val="0"/>\n',
        '          <c:showVal val="0"/>\n',
        '          <c:showCatName val="0"/>\n',
        '          <c:showSerName val="0"/>\n',
        '          <c:showPercent val="0"/>\n',
        '          <c:showBubbleSize val="0"/>\n',
        '          <c:showLeaderLines val="1"/>\n',
        "        </c:dLbls>\n",
        '        <c:firstSliceAng val="0"/>\n',
        '        <c:holeSize val="50"/>\n',
        "      </c:doughnutChart>\n",
        "    </c:plotArea>\n",
        "    <c:legend>\n",
        '      <c:legendPos val="r"/>\n',
        "      <c:layout/>\n",
        '      <c:overlay val="0"/>\n',
        "    </c:legend>\n",
        '    <c:plotVisOnly val="1"/>\n',
        '    <c:dispBlanksAs val="gap"/>\n',
        '    <c:showDLblsOverMax val="0"/>\n',
        "  </c:chart>\n",
        "  <c:txPr>\n",
        "    <a:bodyPr/>\n",
        "    <a:lstStyle/>\n",
        "    <a:p>\n",
        "      <a:pPr>\n",
        '        <a:defRPr sz="1800"/>\n',
        "      </a:pPr>\n",
        "      <a:endParaRPr/>\n",
        "    </a:p>\n",
        "  </c:txPr>\n",
        "</c:chartSpace>\n"
      )
    }
  ),

  private = list(
    .explosion_xml = function() {
      if (private$.chart_type == XL_CHART_TYPE$DOUGHNUT_EXPLODED)
        return('          <c:explosion val="25"/>\n')
      ""
    },

    .ser_xml = function() {
      xml <- ""
      for (series in private$.chart_data$to_list()) {
        w   <- .CategorySeriesXmlWriter$new(series)
        xml <- paste0(xml,
          "        <c:ser>\n",
          sprintf('          <c:idx val="%d"/>\n', series$index),
          sprintf('          <c:order val="%d"/>\n', series$index),
          w$tx_xml(),
          private$.explosion_xml(),
          w$cat_xml(),
          w$val_xml(),
          "        </c:ser>\n"
        )
      }
      xml
    }
  )
)


# ============================================================================
# _LineChartXmlWriter
# ============================================================================

.LineChartXmlWriter <- R6::R6Class(
  ".LineChartXmlWriter",
  inherit = .BaseChartXmlWriter,

  active = list(
    xml = function() {
      paste0(
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n",
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
        '  <c:date1904 val="0"/>\n',
        "  <c:chart>\n",
        '    <c:autoTitleDeleted val="0"/>\n',
        "    <c:plotArea>\n",
        "      <c:lineChart>\n",
        private$.grouping_xml(),
        '        <c:varyColors val="0"/>\n',
        private$.ser_xml(),
        '        <c:marker val="1"/>\n',
        '        <c:smooth val="0"/>\n',
        '        <c:axId val="2118791784"/>\n',
        '        <c:axId val="2140495176"/>\n',
        "      </c:lineChart>\n",
        private$.cat_ax_xml(),
        "      <c:valAx>\n",
        '        <c:axId val="2140495176"/>\n',
        "        <c:scaling/>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="l"/>\n',
        "        <c:majorGridlines/>\n",
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="2118791784"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        "      </c:valAx>\n",
        "    </c:plotArea>\n",
        "    <c:legend>\n",
        '      <c:legendPos val="r"/>\n',
        "      <c:layout/>\n",
        '      <c:overlay val="0"/>\n',
        "    </c:legend>\n",
        '    <c:plotVisOnly val="1"/>\n',
        '    <c:dispBlanksAs val="gap"/>\n',
        '    <c:showDLblsOverMax val="0"/>\n',
        "  </c:chart>\n",
        "  <c:txPr>\n",
        "    <a:bodyPr/>\n",
        "    <a:lstStyle/>\n",
        "    <a:p>\n",
        "      <a:pPr>\n",
        '        <a:defRPr sz="1800"/>\n',
        "      </a:pPr>\n",
        '      <a:endParaRPr lang="en-US"/>\n',
        "    </a:p>\n",
        "  </c:txPr>\n",
        "</c:chartSpace>\n"
      )
    }
  ),

  private = list(
    .grouping_xml = function() {
      XL       <- XL_CHART_TYPE
      standard <- c(XL$LINE, XL$LINE_MARKERS)
      stacked  <- c(XL$LINE_STACKED, XL$LINE_MARKERS_STACKED)
      pct      <- c(XL$LINE_STACKED_100, XL$LINE_MARKERS_STACKED_100)
      if (private$.chart_type %in% standard) return('        <c:grouping val="standard"/>\n')
      if (private$.chart_type %in% stacked)  return('        <c:grouping val="stacked"/>\n')
      if (private$.chart_type %in% pct)      return('        <c:grouping val="percentStacked"/>\n')
      stop("unexpected line chart type", call. = FALSE)
    },

    .marker_xml = function() {
      XL         <- XL_CHART_TYPE
      no_markers <- c(XL$LINE, XL$LINE_STACKED, XL$LINE_STACKED_100)
      if (private$.chart_type %in% no_markers)
        return(paste0("          <c:marker>\n",
                      '            <c:symbol val="none"/>\n',
                      "          </c:marker>\n"))
      ""
    },

    .cat_ax_xml = function() {
      cats <- private$.chart_data$categories
      if (cats$are_dates) {
        return(paste0(
          "      <c:dateAx>\n",
          '        <c:axId val="2118791784"/>\n',
          "        <c:scaling>\n",
          '          <c:orientation val="minMax"/>\n',
          "        </c:scaling>\n",
          '        <c:delete val="0"/>\n',
          '        <c:axPos val="b"/>\n',
          sprintf('        <c:numFmt formatCode="%s" sourceLinked="1"/>\n', cats$number_format),
          '        <c:majorTickMark val="out"/>\n',
          '        <c:minorTickMark val="none"/>\n',
          '        <c:tickLblPos val="nextTo"/>\n',
          '        <c:crossAx val="2140495176"/>\n',
          '        <c:crosses val="autoZero"/>\n',
          '        <c:auto val="1"/>\n',
          '        <c:lblOffset val="100"/>\n',
          '        <c:baseTimeUnit val="days"/>\n',
          "      </c:dateAx>\n"
        ))
      }
      paste0(
        "      <c:catAx>\n",
        '        <c:axId val="2118791784"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="b"/>\n',
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="2140495176"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:auto val="1"/>\n',
        '        <c:lblAlgn val="ctr"/>\n',
        '        <c:lblOffset val="100"/>\n',
        '        <c:noMultiLvlLbl val="0"/>\n',
        "      </c:catAx>\n"
      )
    },

    .ser_xml = function() {
      xml <- ""
      for (series in private$.chart_data$to_list()) {
        w   <- .CategorySeriesXmlWriter$new(series)
        xml <- paste0(xml,
          "        <c:ser>\n",
          sprintf('          <c:idx val="%d"/>\n', series$index),
          sprintf('          <c:order val="%d"/>\n', series$index),
          w$tx_xml(),
          private$.marker_xml(),
          w$cat_xml(),
          w$val_xml(),
          '          <c:smooth val="0"/>\n',
          "        </c:ser>\n"
        )
      }
      xml
    }
  )
)


# ============================================================================
# _PieChartXmlWriter
# ============================================================================

.PieChartXmlWriter <- R6::R6Class(
  ".PieChartXmlWriter",
  inherit = .BaseChartXmlWriter,

  active = list(
    xml = function() {
      paste0(
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n",
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
        "  <c:chart>\n",
        '    <c:autoTitleDeleted val="0"/>\n',
        "    <c:plotArea>\n",
        "      <c:pieChart>\n",
        '        <c:varyColors val="1"/>\n',
        private$.ser_xml(),
        "      </c:pieChart>\n",
        "    </c:plotArea>\n",
        '    <c:dispBlanksAs val="gap"/>\n',
        "  </c:chart>\n",
        "  <c:txPr>\n",
        "    <a:bodyPr/>\n",
        "    <a:lstStyle/>\n",
        "    <a:p>\n",
        "      <a:pPr>\n",
        '        <a:defRPr sz="1800"/>\n',
        "      </a:pPr>\n",
        '      <a:endParaRPr lang="en-US"/>\n',
        "    </a:p>\n",
        "  </c:txPr>\n",
        "</c:chartSpace>\n"
      )
    }
  ),

  private = list(
    .explosion_xml = function() {
      if (private$.chart_type == XL_CHART_TYPE$PIE_EXPLODED)
        return('          <c:explosion val="25"/>\n')
      ""
    },

    .ser_xml = function() {
      series <- private$.chart_data$get(1L)
      w      <- .CategorySeriesXmlWriter$new(series)
      paste0(
        "        <c:ser>\n",
        '          <c:idx val="0"/>\n',
        '          <c:order val="0"/>\n',
        w$tx_xml(),
        private$.explosion_xml(),
        w$cat_xml(),
        w$val_xml(),
        "        </c:ser>\n"
      )
    }
  )
)


# ============================================================================
# _RadarChartXmlWriter
# ============================================================================

.RadarChartXmlWriter <- R6::R6Class(
  ".RadarChartXmlWriter",
  inherit = .BaseChartXmlWriter,

  active = list(
    xml = function() {
      paste0(
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n",
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
        '  <c:date1904 val="0"/>\n',
        '  <c:roundedCorners val="0"/>\n',
        '  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">\n',
        '    <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">\n',
        '      <c14:style val="118"/>\n',
        "    </mc:Choice>\n",
        "    <mc:Fallback>\n",
        '      <c:style val="18"/>\n',
        "    </mc:Fallback>\n",
        "  </mc:AlternateContent>\n",
        "  <c:chart>\n",
        '    <c:autoTitleDeleted val="0"/>\n',
        "    <c:plotArea>\n",
        "      <c:layout/>\n",
        "      <c:radarChart>\n",
        sprintf('        <c:radarStyle val="%s"/>\n', private$.radar_style()),
        '        <c:varyColors val="0"/>\n',
        private$.ser_xml(),
        '        <c:axId val="2073612648"/>\n',
        '        <c:axId val="-2112772216"/>\n',
        "      </c:radarChart>\n",
        "      <c:catAx>\n",
        '        <c:axId val="2073612648"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="b"/>\n',
        "        <c:majorGridlines/>\n",
        '        <c:numFmt formatCode="m/d/yy" sourceLinked="1"/>\n',
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="-2112772216"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:auto val="1"/>\n',
        '        <c:lblAlgn val="ctr"/>\n',
        '        <c:lblOffset val="100"/>\n',
        '        <c:noMultiLvlLbl val="0"/>\n',
        "      </c:catAx>\n",
        "      <c:valAx>\n",
        '        <c:axId val="-2112772216"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="l"/>\n',
        "        <c:majorGridlines/>\n",
        '        <c:numFmt formatCode="General" sourceLinked="1"/>\n',
        '        <c:majorTickMark val="cross"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="2073612648"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:crossBetween val="between"/>\n',
        "      </c:valAx>\n",
        "    </c:plotArea>\n",
        '    <c:plotVisOnly val="1"/>\n',
        '    <c:dispBlanksAs val="gap"/>\n',
        '    <c:showDLblsOverMax val="0"/>\n',
        "  </c:chart>\n",
        "  <c:txPr>\n",
        "    <a:bodyPr/>\n",
        "    <a:lstStyle/>\n",
        "    <a:p>\n",
        "      <a:pPr>\n",
        '        <a:defRPr sz="1800"/>\n',
        "      </a:pPr>\n",
        '      <a:endParaRPr lang="en-US"/>\n',
        "    </a:p>\n",
        "  </c:txPr>\n",
        "</c:chartSpace>\n"
      )
    }
  ),

  private = list(
    .radar_style = function() {
      if (private$.chart_type == XL_CHART_TYPE$RADAR_FILLED) return("filled")
      "marker"
    },

    .marker_xml = function() {
      if (private$.chart_type == XL_CHART_TYPE$RADAR)
        return(paste0("          <c:marker>\n",
                      '            <c:symbol val="none"/>\n',
                      "          </c:marker>\n"))
      ""
    },

    .ser_xml = function() {
      xml <- ""
      for (series in private$.chart_data$to_list()) {
        w   <- .CategorySeriesXmlWriter$new(series)
        xml <- paste0(xml,
          "        <c:ser>\n",
          sprintf('          <c:idx val="%d"/>\n', series$index),
          sprintf('          <c:order val="%d"/>\n', series$index),
          w$tx_xml(),
          private$.marker_xml(),
          w$cat_xml(),
          w$val_xml(),
          '          <c:smooth val="0"/>\n',
          "        </c:ser>\n"
        )
      }
      xml
    }
  )
)


# ============================================================================
# _XyChartXmlWriter (scatter)
# ============================================================================

.XyChartXmlWriter <- R6::R6Class(
  ".XyChartXmlWriter",
  inherit = .BaseChartXmlWriter,

  active = list(
    xml = function() {
      paste0(
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n",
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
        "  <c:chart>\n",
        "    <c:plotArea>\n",
        "      <c:scatterChart>\n",
        sprintf('        <c:scatterStyle val="%s"/>\n', private$.scatterStyle_val()),
        '        <c:varyColors val="0"/>\n',
        private$.ser_xml(),
        '        <c:axId val="-2128940872"/>\n',
        '        <c:axId val="-2129643912"/>\n',
        "      </c:scatterChart>\n",
        "      <c:valAx>\n",
        '        <c:axId val="-2128940872"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="b"/>\n',
        '        <c:numFmt formatCode="General" sourceLinked="1"/>\n',
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="-2129643912"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:crossBetween val="midCat"/>\n',
        "      </c:valAx>\n",
        "      <c:valAx>\n",
        '        <c:axId val="-2129643912"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="l"/>\n',
        "        <c:majorGridlines/>\n",
        '        <c:numFmt formatCode="General" sourceLinked="1"/>\n',
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="-2128940872"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:crossBetween val="midCat"/>\n',
        "      </c:valAx>\n",
        "    </c:plotArea>\n",
        "    <c:legend>\n",
        '      <c:legendPos val="r"/>\n',
        "      <c:layout/>\n",
        '      <c:overlay val="0"/>\n',
        "    </c:legend>\n",
        '    <c:plotVisOnly val="1"/>\n',
        '    <c:dispBlanksAs val="gap"/>\n',
        '    <c:showDLblsOverMax val="0"/>\n',
        "  </c:chart>\n",
        "  <c:txPr>\n",
        "    <a:bodyPr/>\n",
        "    <a:lstStyle/>\n",
        "    <a:p>\n",
        "      <a:pPr>\n",
        '        <a:defRPr sz="1800"/>\n',
        "      </a:pPr>\n",
        '      <a:endParaRPr lang="en-US"/>\n',
        "    </a:p>\n",
        "  </c:txPr>\n",
        "</c:chartSpace>\n"
      )
    }
  ),

  private = list(
    .scatterStyle_val = function() {
      XL     <- XL_CHART_TYPE
      smooth <- c(XL$XY_SCATTER_SMOOTH, XL$XY_SCATTER_SMOOTH_NO_MARKERS)
      if (private$.chart_type %in% smooth) return("smoothMarker")
      "lineMarker"
    },

    .marker_xml = function() {
      XL         <- XL_CHART_TYPE
      no_markers <- c(XL$XY_SCATTER_LINES_NO_MARKERS, XL$XY_SCATTER_SMOOTH_NO_MARKERS)
      if (private$.chart_type %in% no_markers)
        return(paste0("          <c:marker>\n",
                      '            <c:symbol val="none"/>\n',
                      "          </c:marker>\n"))
      ""
    },

    .spPr_xml = function() {
      if (private$.chart_type == XL_CHART_TYPE$XY_SCATTER)
        return(paste0(
          "          <c:spPr>\n",
          '            <a:ln w="47625">\n',
          "              <a:noFill/>\n",
          "            </a:ln>\n",
          "          </c:spPr>\n"
        ))
      ""
    },

    .ser_xml = function() {
      xml <- ""
      for (series in private$.chart_data$to_list()) {
        w   <- .XySeriesXmlWriter$new(series)
        xml <- paste0(xml,
          "        <c:ser>\n",
          sprintf('          <c:idx val="%d"/>\n', series$index),
          sprintf('          <c:order val="%d"/>\n', series$index),
          w$tx_xml(),
          private$.spPr_xml(),
          private$.marker_xml(),
          w$xVal_xml(),
          w$yVal_xml(),
          '          <c:smooth val="0"/>\n',
          "        </c:ser>\n"
        )
      }
      xml
    }
  )
)


# ============================================================================
# _BubbleChartXmlWriter
# ============================================================================

.BubbleChartXmlWriter <- R6::R6Class(
  ".BubbleChartXmlWriter",
  inherit = .XyChartXmlWriter,

  active = list(
    xml = function() {
      paste0(
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n",
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
        "  <c:chart>\n",
        '    <c:autoTitleDeleted val="0"/>\n',
        "    <c:plotArea>\n",
        "      <c:layout/>\n",
        "      <c:bubbleChart>\n",
        '        <c:varyColors val="0"/>\n',
        private$.ser_xml(),
        "        <c:dLbls>\n",
        '          <c:showLegendKey val="0"/>\n',
        '          <c:showVal val="0"/>\n',
        '          <c:showCatName val="0"/>\n',
        '          <c:showSerName val="0"/>\n',
        '          <c:showPercent val="0"/>\n',
        '          <c:showBubbleSize val="0"/>\n',
        "        </c:dLbls>\n",
        '        <c:bubbleScale val="100"/>\n',
        '        <c:showNegBubbles val="0"/>\n',
        '        <c:axId val="-2115720072"/>\n',
        '        <c:axId val="-2115723560"/>\n',
        "      </c:bubbleChart>\n",
        "      <c:valAx>\n",
        '        <c:axId val="-2115720072"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="b"/>\n',
        '        <c:numFmt formatCode="General" sourceLinked="1"/>\n',
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="-2115723560"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:crossBetween val="midCat"/>\n',
        "      </c:valAx>\n",
        "      <c:valAx>\n",
        '        <c:axId val="-2115723560"/>\n',
        "        <c:scaling>\n",
        '          <c:orientation val="minMax"/>\n',
        "        </c:scaling>\n",
        '        <c:delete val="0"/>\n',
        '        <c:axPos val="l"/>\n',
        "        <c:majorGridlines/>\n",
        '        <c:numFmt formatCode="General" sourceLinked="1"/>\n',
        '        <c:majorTickMark val="out"/>\n',
        '        <c:minorTickMark val="none"/>\n',
        '        <c:tickLblPos val="nextTo"/>\n',
        '        <c:crossAx val="-2115720072"/>\n',
        '        <c:crosses val="autoZero"/>\n',
        '        <c:crossBetween val="midCat"/>\n',
        "      </c:valAx>\n",
        "    </c:plotArea>\n",
        "    <c:legend>\n",
        '      <c:legendPos val="r"/>\n',
        "      <c:layout/>\n",
        '      <c:overlay val="0"/>\n',
        "    </c:legend>\n",
        '    <c:plotVisOnly val="1"/>\n',
        '    <c:dispBlanksAs val="gap"/>\n',
        '    <c:showDLblsOverMax val="0"/>\n',
        "  </c:chart>\n",
        "  <c:txPr>\n",
        "    <a:bodyPr/>\n",
        "    <a:lstStyle/>\n",
        "    <a:p>\n",
        "      <a:pPr>\n",
        '        <a:defRPr sz="1800"/>\n',
        "      </a:pPr>\n",
        '      <a:endParaRPr lang="en-US"/>\n',
        "    </a:p>\n",
        "  </c:txPr>\n",
        "</c:chartSpace>\n"
      )
    }
  ),

  private = list(
    .bubble3D_val = function() {
      if (private$.chart_type == XL_CHART_TYPE$BUBBLE_THREE_D_EFFECT) return("1")
      "0"
    },

    .ser_xml = function() {
      xml <- ""
      for (series in private$.chart_data$to_list()) {
        w   <- .BubbleSeriesXmlWriter$new(series)
        xml <- paste0(xml,
          "        <c:ser>\n",
          sprintf('          <c:idx val="%d"/>\n', series$index),
          sprintf('          <c:order val="%d"/>\n', series$index),
          w$tx_xml(),
          '          <c:invertIfNegative val="0"/>\n',
          w$xVal_xml(),
          w$yVal_xml(),
          w$bubbleSize_xml(),
          sprintf('          <c:bubble3D val="%s"/>\n', private$.bubble3D_val()),
          "        </c:ser>\n"
        )
      }
      xml
    }
  )
)
