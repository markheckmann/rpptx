# Chart part objects.
#
# Ported from python-pptx/src/pptx/parts/chart.py and
# python-pptx/src/pptx/parts/embeddedpackage.py.


# ============================================================================
# EmbeddedXlsxPart — embedded Excel workbook
# ============================================================================

#' Embedded Excel workbook part
#'
#' Stores the raw xlsx bytes for a chart's data workbook.
#'
#' @keywords internal
#' @export
EmbeddedXlsxPart <- R6::R6Class(
  "EmbeddedXlsxPart",
  inherit = Part,

  public = list(
    load = function(partname, content_type, package, blob) {
      EmbeddedXlsxPart$new(partname, content_type, package, blob)
    }
  )
)

#' Create a new EmbeddedXlsxPart containing `blob`
#' @param blob Raw bytes (the xlsx file).
#' @param package The OPC package.
#' @return An EmbeddedXlsxPart.
#' @keywords internal
#' @export
EmbeddedXlsxPart_new <- function(blob, package) {
  partname <- package$next_partname("/ppt/embeddings/Microsoft_Excel_Sheet%d.xlsx")
  EmbeddedXlsxPart$new(partname, CT$SML_SHEET, package, blob)
}


# ============================================================================
# ChartWorkbook — links a ChartPart to its xlsx data
# ============================================================================

#' Manages the embedded Excel workbook for a chart
#' @keywords internal
#' @export
ChartWorkbook <- R6::R6Class(
  "ChartWorkbook",

  public = list(
    initialize = function(chart_space_elm, chart_part) {
      private$.chart_space <- chart_space_elm
      private$.chart_part  <- chart_part
    },

    # Replace (or create) the embedded xlsx part with new data.
    update_from_xlsx_blob = function(xlsx_blob) {
      xlsx_part <- private$.xlsx_part()
      if (is.null(xlsx_part)) {
        new_part <- EmbeddedXlsxPart_new(xlsx_blob, private$.chart_part$package)
        private$.set_xlsx_part(new_part)
      } else {
        xlsx_part$blob <- xlsx_blob
      }
    }
  ),

  private = list(
    .chart_space = NULL,
    .chart_part  = NULL,

    # Return the current EmbeddedXlsxPart, or NULL if none linked yet.
    .xlsx_part = function() {
      rId <- private$.xlsx_part_rId()
      if (is.null(rId)) return(NULL)
      private$.chart_part$related_part(rId)
    },

    # Add a relationship to `xlsx_part` and record it in <c:externalData r:id=>.
    .set_xlsx_part = function(xlsx_part) {
      rId <- private$.chart_part$relate_to(xlsx_part, RT$PACKAGE)
      private$.add_external_data(rId)
    },

    # Return the r:id of the existing <c:externalData>, or NULL.
    .xlsx_part_rId = function() {
      node   <- private$.chart_space$get_node()
      c_ns   <- .nsmap[["c"]]
      r_ns   <- .nsmap[["r"]]
      ext_nd <- xml2::xml_find_first(
        node, "c:externalData",
        ns = c(c = c_ns, r = r_ns)
      )
      if (inherits(ext_nd, "xml_missing")) return(NULL)
      val <- xml2::xml_attr(ext_nd, "r:id",
                            ns = c(r = r_ns))
      if (is.na(val)) NULL else val
    },

    # Append <c:externalData r:id="rId"/> to the chartSpace element.
    .add_external_data = function(rId) {
      c_ns <- .nsmap[["c"]]
      r_ns <- .nsmap[["r"]]
      xml_str <- sprintf(
        '<c:externalData xmlns:c="%s" xmlns:r="%s" r:id="%s"/>',
        c_ns, r_ns, rId
      )
      ext_node <- xml2::xml_root(xml2::read_xml(xml_str))
      private$.chart_space$append_child(ext_node)
    }
  )
)


# ============================================================================
# ChartPart — /ppt/charts/chartN.xml
# ============================================================================

#' Chart XML part
#'
#' Wraps a c:chartSpace XML document. Created via [ChartPart_new()].
#'
#' @keywords internal
#' @export
ChartPart <- R6::R6Class(
  "ChartPart",
  inherit = XmlPart,

  public = list(
    # Class-method-style loader (called by PartFactory_create).
    load = function(partname, content_type, package, blob) {
      element <- rpptx_parse_xml(blob)
      ChartPart$new(partname, content_type, package, element)
    }
  ),

  active = list(
    # Chart domain object for this chart part.
    chart = function() {
      Chart$new(self$element, self)
    },

    # ChartWorkbook that manages the embedded xlsx.
    chart_workbook = function() {
      if (is.null(private$.workbook_cache)) {
        private$.workbook_cache <- ChartWorkbook$new(self$element, self)
      }
      private$.workbook_cache
    }
  ),

  private = list(
    .workbook_cache = NULL
  )
)

#' Create a new ChartPart containing `chart_type` data from `chart_data`
#'
#' @param chart_type An integer from XL_CHART_TYPE.
#' @param chart_data A CategoryChartData, XyChartData, or BubbleChartData.
#' @param package The OPC package.
#' @return A ChartPart.
#' @keywords internal
#' @export
ChartPart_new <- function(chart_type, chart_data, package) {
  partname  <- package$next_partname("/ppt/charts/chart%d.xml")
  xml_str   <- chart_xml_writer(chart_type, chart_data)$xml
  xml_bytes <- charToRaw(xml_str)
  chart_part <- XmlPart_load(ChartPart, partname, CT$DML_CHART, package, xml_bytes)
  chart_part$chart_workbook$update_from_xlsx_blob(chart_data$xlsx_blob)
  chart_part
}


# ============================================================================
# .onLoad registration
# ============================================================================

.onLoad_parts_chart <- function() {
  register_part_type(CT$DML_CHART, ChartPart)
}
