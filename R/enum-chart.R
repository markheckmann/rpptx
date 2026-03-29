# Chart-related enumeration constants.
#
# Ported from python-pptx/src/pptx/enum/chart.py.


#' XL_CHART_TYPE — chart type enumeration
#'
#' Integer constants specifying chart types, mirroring XlChartType.
#'
#' @noRd
#' @export
XL_CHART_TYPE <- list(
  THREE_D_AREA                 =  -4098L,
  THREE_D_AREA_STACKED         =    78L,
  THREE_D_AREA_STACKED_100     =    79L,
  THREE_D_BAR_CLUSTERED        =    60L,
  THREE_D_BAR_STACKED          =    61L,
  THREE_D_BAR_STACKED_100      =    62L,
  THREE_D_COLUMN               = -4100L,
  THREE_D_COLUMN_CLUSTERED     =    54L,
  THREE_D_COLUMN_STACKED       =    55L,
  THREE_D_COLUMN_STACKED_100   =    56L,
  THREE_D_LINE                 = -4101L,
  THREE_D_PIE                  = -4102L,
  THREE_D_PIE_EXPLODED         =    70L,
  AREA                         =     1L,
  AREA_STACKED                 =    76L,
  AREA_STACKED_100             =    77L,
  BAR_CLUSTERED                =    57L,
  BAR_OF_PIE                   =    71L,
  BAR_STACKED                  =    58L,
  BAR_STACKED_100              =    59L,
  BUBBLE                       =    15L,
  BUBBLE_THREE_D_EFFECT        =    87L,
  COLUMN_CLUSTERED             =    51L,
  COLUMN_STACKED               =    52L,
  COLUMN_STACKED_100           =    53L,
  CONE_BAR_CLUSTERED           =   102L,
  CONE_BAR_STACKED             =   103L,
  CONE_BAR_STACKED_100         =   104L,
  CONE_COL                     =   105L,
  CONE_COL_CLUSTERED           =    99L,
  CONE_COL_STACKED             =   100L,
  CONE_COL_STACKED_100         =   101L,
  CYLINDER_BAR_CLUSTERED       =    95L,
  CYLINDER_BAR_STACKED         =    96L,
  CYLINDER_BAR_STACKED_100     =    97L,
  CYLINDER_COL                 =    98L,
  CYLINDER_COL_CLUSTERED       =    92L,
  CYLINDER_COL_STACKED         =    93L,
  CYLINDER_COL_STACKED_100     =    94L,
  DOUGHNUT                     = -4120L,
  DOUGHNUT_EXPLODED            =    80L,
  LINE                         =     4L,
  LINE_MARKERS                 =    65L,
  LINE_MARKERS_STACKED         =    66L,
  LINE_MARKERS_STACKED_100     =    67L,
  LINE_STACKED                 =    63L,
  LINE_STACKED_100             =    64L,
  PIE                          =     5L,
  PIE_EXPLODED                 =    69L,
  PIE_OF_PIE                   =    68L,
  PYRAMID_BAR_CLUSTERED        =   109L,
  PYRAMID_BAR_STACKED          =   110L,
  PYRAMID_BAR_STACKED_100      =   111L,
  PYRAMID_COL                  =   112L,
  PYRAMID_COL_CLUSTERED        =   106L,
  PYRAMID_COL_STACKED          =   107L,
  PYRAMID_COL_STACKED_100      =   108L,
  RADAR                        = -4151L,
  RADAR_FILLED                 =    82L,
  RADAR_MARKERS                =    81L,
  STOCK_HLC                    =    88L,
  STOCK_OHLC                   =    89L,
  STOCK_VHLC                   =    90L,
  STOCK_VOHLC                  =    91L,
  SURFACE                      =    83L,
  SURFACE_TOP_VIEW             =    85L,
  SURFACE_TOP_VIEW_WIREFRAME   =    86L,
  SURFACE_WIREFRAME            =    84L,
  XY_SCATTER                   = -4169L,
  XY_SCATTER_LINES             =    74L,
  XY_SCATTER_LINES_NO_MARKERS  =    75L,
  XY_SCATTER_SMOOTH            =    72L,
  XY_SCATTER_SMOOTH_NO_MARKERS =    73L
)


#' XL_AXIS_CROSSES — where the other axis crosses this axis
#'
#' XML string values for `c:crosses/@val`. Use with axis `crosses` property.
#'
#' @noRd
#' @export
XL_AXIS_CROSSES <- list(
  AUTOMATIC = "autoZero",
  CUSTOM    = "",
  MAXIMUM   = "max",
  MINIMUM   = "min"
)


#' XL_CATEGORY_TYPE — category axis scale type
#'
#' Integer constants (read-only identifiers, not XML-encoded directly).
#'
#' @noRd
#' @export
XL_CATEGORY_TYPE <- list(
  AUTOMATIC_SCALE =  -4105L,
  CATEGORY_SCALE  =     2L,
  TIME_SCALE      =     3L
)


#' XL_DATA_LABEL_POSITION — position of a data label
#'
#' XML string values for `c:dLblPos/@val`.
#'
#' @noRd
#' @export
XL_DATA_LABEL_POSITION <- list(
  ABOVE       = "t",
  BELOW       = "b",
  BEST_FIT    = "bestFit",
  CENTER      = "ctr",
  INSIDE_BASE = "inBase",
  INSIDE_END  = "inEnd",
  LEFT        = "l",
  MIXED       = "",
  OUTSIDE_END = "outEnd",
  RIGHT       = "r"
)

#' XL_LABEL_POSITION — alias for XL_DATA_LABEL_POSITION
#' @noRd
#' @export
XL_LABEL_POSITION <- XL_DATA_LABEL_POSITION


#' XL_LEGEND_POSITION — position of chart legend
#'
#' XML string values for `c:legendPos/@val`.
#'
#' @noRd
#' @export
XL_LEGEND_POSITION <- list(
  BOTTOM = "b",
  CORNER = "tr",
  CUSTOM = "",
  LEFT   = "l",
  RIGHT  = "r",
  TOP    = "t"
)


#' XL_MARKER_STYLE — shape of a data point marker
#'
#' XML string values for `c:symbol/@val`.
#'
#' @noRd
#' @export
XL_MARKER_STYLE <- list(
  AUTOMATIC = "auto",
  CIRCLE    = "circle",
  DASH      = "dash",
  DIAMOND   = "diamond",
  DOT       = "dot",
  NONE      = "none",
  PICTURE   = "picture",
  PLUS      = "plus",
  SQUARE    = "square",
  STAR      = "star",
  TRIANGLE  = "triangle",
  X         = "x"
)


#' XL_TICK_LABEL_POSITION — position of tick-mark labels on a chart axis
#'
#' XML string values for `c:tickLblPos/@val`.
#'
#' @noRd
#' @export
XL_TICK_LABEL_POSITION <- list(
  HIGH         = "high",
  LOW          = "low",
  NEXT_TO_AXIS = "nextTo",
  NONE         = "none"
)


#' XL_TICK_MARK — type of axis tick mark
#'
#' XML string values for `c:majorTickMark/@val` and `c:minorTickMark/@val`.
#'
#' @noRd
#' @export
XL_TICK_MARK <- list(
  CROSS   = "cross",
  INSIDE  = "in",
  NONE    = "none",
  OUTSIDE = "out"
)
