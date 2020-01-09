//============================================================================
// MapInfo version 17.0.3 - System defines
//----------------------------------------------------------------------------
// This file contains defines useful when programming in the MapBasic
// language.  There are three versions of this file:
//     MAPBASIC.DEF  - MapBasic syntax
//     MAPBASIC.BAS  - Visual Basic syntax
//     MAPBASIC.H    - C/C++ syntax
//----------------------------------------------------------------------------
// The defines in this file are organized into the following sections:
//     General Purpose defines:
//       macros, logical constants, angle conversion, colors, string length
//     BrowserInfo() defines
//     ButtonPadInfo() defines
//     ColumnInfo() and column type defines
//     CommandInfo() and task switch defines
//     DateWindow() defines
//     FileAttr() and file access mode defines
//     GetFolderPath$() defines
//     GetPreference() defines - See MapBasic.DEF file for comments
//     GetPreferencePath$() defines
//     IntersectNodes() parameters
//     LabelInfo() defines
//     GroupLayerInfo() defines
//     LayerListInfo() defines
//     LayerInfo(), display mode, label property, layer type, hotlink defines
//     LayoutInfo(), LayoutItemInfo() and LayoutPageItemInfo() defines
//     LegendInfo() and legend orientation defines
//     LegendFrameInfo() and frame type defines
//     LegendStyleInfo() defines
//     LibraryServiceInfo() defines
//     LocateFile$() defines
//     Map3DInfo() defines
//     MapperInfo(), display mode, calculation type, and clip type defines
//     MenuItemInfoByID() and MenuItemInfoByHandler() defines
//     ObjectGeography() defines
//     ObjectInfo() and object type defines
//     PrismMapInfo() defines
//     SearchInfo() defines
//     SelectionInfo() defines
//     Server statement and function defines
//     SessionInfo() defines
//     Set Next Document Style defines
//     StringCompare() return values
//     StyleAttr() defines
//     SystemInfo(), platform, and version defines
//     TableInfo() and table type defines
//     UndoInfo() defines
//     ThemeInfo() defines
//     ThemeItemInfo() defines
//     ThemeExpressionInfo() defines
//     RangeThemeInfo() defines
//     RangeThemeItemInfo() defines
//     IndividualThemeInfo() defines
//     IndividualThemeItemInfo() defines
//     DotDensityThemeInfo() defines
//     GraduatedThemeInfo() defines
//     BarChartThemeInfo() defines
//     BarChartThemeItemInfo() defines
//     PieChartThemeInfo() defines
//     PieChartThemeItemInfo() defines
//     WindowInfo(), window type and state, and print orientation defines
//     Abbreviated list of error codes
//     Backward Compatibility defines
//============================================================================
// MAPBASIC.DEF is converted into MAPBASIC.H by doing the following:
//   - concatenate MAPBASIC.DEF and MENU.DEF into MAPBASIC.H
//   - search & replace "'" at begining of a line with "//"
//   - search & replace "Define" at begining of a line with "#define"
//   - delete the following sections:
//         * General Purpose defines:
//              Macros, Logical Constants, Angle Conversions
//         * Abbreviated list of error codes
//         * Backward Compatibility defines
//         * Menu constants whose names have changed
//         * Obsolete menu items
//============================================================================
// MAPBASIC.DEF is converted into MAPBASIC.BAS by doing the following:
//   - concatenate MAPBASIC.DEF and MENU.DEF into MAPBASIC.BAS
//   - search & replace "Define <name>" with "Global Const <name> ="
//     e.g. "<Define {[!-z]+} +{[!-z]}" with "Global Const \0 = \1" with Brief
//   - delete the following sections:
//         * General Purpose defines:
//             Macros, Logical Constants, Angle Conversions
//         * Abbreviated list of error codes
//         * Backward Compatibility defines
//         * Menu constants whose names have changed
//         * Obsolete menu items
//============================================================================

//============================================================================
// General Purpose defines
//============================================================================
//----------------------------------------------------------------------------
// Colors
//----------------------------------------------------------------------------
#define BLACK                                    0
#define WHITE                                    16777215
#define RED                                      16711680
#define GREEN                                    65280
#define BLUE                                     255
#define CYAN                                     65535
#define MAGENTA                                  16711935
#define YELLOW                                   16776960

//----------------------------------------------------------------------------
//Maximum length for character string
//----------------------------------------------------------------------------
#define MAX_STRING_LENGTH                        32767

//============================================================================
// BrowserInfo() defines
//============================================================================
#define BROWSER_INFO_NROWS                       1
#define BROWSER_INFO_NCOLS                       2
#define BROWSER_INFO_CURRENT_ROW                 3
#define BROWSER_INFO_CURRENT_COLUMN              4
#define BROWSER_INFO_CURRENT_CELL_VALUE          5

//============================================================================
// ButtonPadInfo() defines
//============================================================================
#define BTNPAD_INFO_FLOATING                     1
#define BTNPAD_INFO_WIDTH                        2
#define BTNPAD_INFO_NBTNS                        3
#define BTNPAD_INFO_X                            4
#define BTNPAD_INFO_Y                            5
#define BTNPAD_INFO_WINID                        6
#define BTNPAD_INFO_DOCK_POSITION                7 // New as per MI Pro 10.5 to support docking left, right and bottom

/*-----------------------------------------------------------------------*
| New as per MI Pro 10.5.
| Codes returned from ButtonPadInfo() when 'BTNPAD_INFO_DOCK_POSITION' code
| is used to inquiry about the tool bar position
*------------------------------------------------------------------------*/
#define BTNPAD_INFO_DOCK_NONE                    0 // Floating position
#define BTNPAD_INFO_DOCK_LEFT                    1 // Left position
#define BTNPAD_INFO_DOCK_TOP                     2 // Top position
#define BTNPAD_INFO_DOCK_RIGHT                   3 // Right position
#define BTNPAD_INFO_DOCK_BOTTOM                  4 // Bottom position

//============================================================================
// ColumnInfo() defines
//============================================================================
#define COL_INFO_NAME                            1
#define COL_INFO_NUM                             2
#define COL_INFO_TYPE                            3
#define COL_INFO_WIDTH                           4
#define COL_INFO_DECPLACES                       5
#define COL_INFO_INDEXED                         6
#define COL_INFO_EDITABLE                        7

//----------------------------------------------------------------------------
// Column type defines, returned by ColumnInfo() for COL_INFO_TYPE
//----------------------------------------------------------------------------
#define COL_TYPE_CHAR                            1
#define COL_TYPE_DECIMAL                         2
#define COL_TYPE_INTEGER                         3
#define COL_TYPE_SMALLINT                        4
#define COL_TYPE_DATE                            5
#define COL_TYPE_LOGICAL                         6
#define COL_TYPE_GRAPHIC                         7
#define COL_TYPE_FLOAT                           8
#define COL_TYPE_TIME                            37
#define COL_TYPE_DATETIME                        38
#define COL_TYPE_LARGEINT                        39

//============================================================================
// CommandInfo() defines
//============================================================================
#define CMD_INFO_X                               1
#define CMD_INFO_Y                               2
#define CMD_INFO_SHIFT                           3
#define CMD_INFO_CTRL                            4
#define CMD_INFO_X2                              5
#define CMD_INFO_Y2                              6
#define CMD_INFO_TOOLBTN                         7
#define CMD_INFO_MENUITEM                        8
#define CMD_INFO_WIN                             1
#define CMD_INFO_SELTYPE                         1
#define CMD_INFO_ROWID                           2
#define CMD_INFO_INTERRUPT                       3
#define CMD_INFO_STATUS                          1
#define CMD_INFO_MSG                             1000
#define CMD_INFO_DLG_OK                          1
#define CMD_INFO_DLG_DBL                         1
#define CMD_INFO_FIND_RC                         3
#define CMD_INFO_FIND_ROWID                      4
#define CMD_INFO_XCMD                            1
#define CMD_INFO_CUSTOM_OBJ                      1
#define CMD_INFO_TASK_SWITCH                     1
#define CMD_INFO_EDIT_TABLE                      1
#define CMD_INFO_EDIT_STATUS                     2
#define CMD_INFO_EDIT_ASK                        1
#define CMD_INFO_EDIT_SAVE                       2
#define CMD_INFO_EDIT_DISCARD                    3
#define CMD_INFO_HL_WINDOW_ID                    17
#define CMD_INFO_HL_TABLE_NAME                   18
#define CMD_INFO_HL_ROWID                        19
#define CMD_INFO_HL_LAYER_ID                     20
#define CMD_INFO_HL_FILE_NAME                    21

//----------------------------------------------------------------------------
// Task Switches, returned by CommandInfo() for CMD_INFO_TASK_SWITCH
//----------------------------------------------------------------------------
#define SWITCHING_OUT_OF_MAPINFO                 0
#define SWITCHING_INTO_MAPINFO                   1

//============================================================================
// DateWindow() defines
//============================================================================
#define DATE_WIN_SESSION                         1
#define DATE_WIN_CURPROG                         2

//============================================================================
// FileAttr() defines
//============================================================================
#define FILE_ATTR_MODE                           1
#define FILE_ATTR_FILESIZE                       2

//----------------------------------------------------------------------------
// File Access Modes, returned by FileAttr() for FILE_ATTR_MODE
//----------------------------------------------------------------------------
#define MODE_INPUT                               0
#define MODE_OUTPUT                              1
#define MODE_APPEND                              2
#define MODE_RANDOM                              3
#define MODE_BINARY                              4

//============================================================================
// GetFolderPath$() defines
//============================================================================

#define FOLDER_MI_APPDATA                        -1
#define FOLDER_MI_LOCAL_APPDATA                  -2
#define FOLDER_MI_PREFERENCE                     -3
#define FOLDER_MI_COMMON_APPDATA                 -4
#define FOLDER_APPDATA                           26
#define FOLDER_LOCAL_APPDATA                     28
#define FOLDER_COMMON_APPDATA                    35
#define FOLDER_COMMON_DOCS                       46
#define FOLDER_MYDOCS                            5
#define FOLDER_MYPICS                            39

//============================================================================
// GetPreferencePath$(), GetCurrentPath$(), and Set Path defines
//============================================================================

#define PREFERENCE_PATH_TABLE          0
#define PREFERENCE_PATH_WORKSPACE      1
#define PREFERENCE_PATH_MBX            2
#define PREFERENCE_PATH_IMPORT         3
#define PREFERENCE_PATH_SQLQUERY       4
#define PREFERENCE_PATH_THEMETHEMPLATE 5
#define PREFERENCE_PATH_MIQUERY        6
#define PREFERENCE_PATH_NEWGRID        7
#define PREFERENCE_PATH_CRYSTAL        8
#define PREFERENCE_PATH_GRAPHSUPPORT   9
#define PREFERENCE_PATH_REMOTETABLE    10
#define PREFERENCE_PATH_SHAPEFILE      11
#define PREFERENCE_PATH_WFSTABLE       12
#define PREFERENCE_PATH_WMSTABLE       13
#define PREFERENCE_PATH_LAYOUTIMAGES   14

//============================================================================
// IntersectNodes() defines
//============================================================================
#define INCL_CROSSINGS                           1
#define INCL_COMMON                              6
#define INCL_ALL                                 7

//============================================================================
// LabelInfo() defines
//============================================================================
#define LABEL_INFO_OBJECT                        1
#define LABEL_INFO_POSITION                      2
#define LABEL_INFO_ANCHORX                       3
#define LABEL_INFO_ANCHORY                       4
#define LABEL_INFO_OFFSET                        5
#define LABEL_INFO_ROWID                         6
#define LABEL_INFO_TABLE                         7
#define LABEL_INFO_EDIT                          8
#define LABEL_INFO_EDIT_VISIBILITY               9
#define LABEL_INFO_EDIT_ANCHOR                   10
#define LABEL_INFO_EDIT_OFFSET                   11
#define LABEL_INFO_EDIT_FONT                     12
#define LABEL_INFO_EDIT_PEN                      13
#define LABEL_INFO_EDIT_TEXT                     14
#define LABEL_INFO_EDIT_TEXTARROW                15
#define LABEL_INFO_EDIT_ANGLE                    16
#define LABEL_INFO_EDIT_POSITION                 17
#define LABEL_INFO_EDIT_TEXTLINE                 18
#define LABEL_INFO_SELECT                        19
#define LABEL_INFO_DRAWN                         20
#define LABEL_INFO_ORIENTATION                   21

//============================================================================
// Codes passed to the GroupLayerInfo function to get info about a group layer.
//============================================================================
#define GROUPLAYER_INFO_NAME                     1
#define GROUPLAYER_INFO_LAYERLIST_ID             2
#define GROUPLAYER_INFO_DISPLAY                  3
#define GROUPLAYER_INFO_LAYERS                   4
#define GROUPLAYER_INFO_ALL_LAYERS               5
#define GROUPLAYER_INFO_TOPLEVEL_LAYERS          6
#define GROUPLAYER_INFO_PARENT_GROUP_ID          7

//============================================================================
// Values returned by GroupLayerInfo() for GROUPLAYER_INFO_DISPLAY. These
// defines correspond to the MapBasic defines in MAPBASIC.DEF. If you alter
// these you must alter MAPBASIC.DEF.
//============================================================================
#define GROUPLAYER_INFO_DISPLAY_OFF              0
#define GROUPLAYER_INFO_DISPLAY_ON               1

/****************************************************************************/
/* Codes passed to the LayerListInfo function to help enumerating all layers in a Map. */
/****************************************************************************/
#define LAYERLIST_INFO_TYPE              1
#define LAYERLIST_INFO_NAME              2
#define LAYERLIST_INFO_LAYER_ID          3
#define LAYERLIST_INFO_GROUPLAYER_ID     4

/****************************************************************************/
/* Values returned by LayerListInfo() for LAYERLIST_INFO_TYPE. These        */
/* defines correspond to the MapBasic defines in MAPBASIC.DEF. If you alter */
/* these you must alter MAPBASIC.DEF.                                       */
/****************************************************************************/
#define LAYERLIST_INFO_TYPE_LAYER        0
#define LAYERLIST_INFO_TYPE_GROUP        1

//============================================================================
// LayerInfo() defines
//============================================================================
#define LAYER_INFO_NAME                          1
#define LAYER_INFO_EDITABLE                      2
#define LAYER_INFO_SELECTABLE                    3
#define LAYER_INFO_ZOOM_LAYERED                  4
#define LAYER_INFO_ZOOM_MIN                      5
#define LAYER_INFO_ZOOM_MAX                      6
#define LAYER_INFO_COSMETIC                      7
#define LAYER_INFO_PATH                          8
#define LAYER_INFO_DISPLAY                       9
#define LAYER_INFO_OVR_LINE                      10
#define LAYER_INFO_OVR_PEN                       11
#define LAYER_INFO_OVR_BRUSH                     12
#define LAYER_INFO_OVR_SYMBOL                    13
#define LAYER_INFO_OVR_FONT                      14
#define LAYER_INFO_LBL_EXPR                      15
#define LAYER_INFO_LBL_LT                        16
#define LAYER_INFO_LBL_CURFONT                   17
#define LAYER_INFO_LBL_FONT                      18
#define LAYER_INFO_LBL_PARALLEL                  19
#define LAYER_INFO_LBL_POS                       20
#define LAYER_INFO_ARROWS                        21
#define LAYER_INFO_NODES                         22
#define LAYER_INFO_CENTROIDS                     23
#define LAYER_INFO_TYPE                          24
#define LAYER_INFO_LBL_VISIBILITY                25
#define LAYER_INFO_LBL_ZOOM_MIN                  26
#define LAYER_INFO_LBL_ZOOM_MAX                  27
#define LAYER_INFO_LBL_AUTODISPLAY               28
#define LAYER_INFO_LBL_OVERLAP                   29
#define LAYER_INFO_LBL_DUPLICATES                30
#define LAYER_INFO_LBL_OFFSET                    31
#define LAYER_INFO_LBL_MAX                       32
#define LAYER_INFO_LBL_PARTIALSEGS               33
#define LAYER_INFO_HOTLINK_EXPR                  34
#define LAYER_INFO_HOTLINK_MODE                  35
#define LAYER_INFO_HOTLINK_RELATIVE              36
#define LAYER_INFO_HOTLINK_COUNT                 37
#define LAYER_INFO_LBL_ORIENTATION               38
#define LAYER_INFO_LAYER_ALPHA                   39
#define LAYER_INFO_LAYER_TRANSLUCENCY            40
#define LAYER_INFO_LABEL_ALPHA                   41
#define LAYER_INFO_LAYERLIST_ID                  42
#define LAYER_INFO_PARENT_GROUP_ID               43
// Code 44 - 52 are for override style & label
#define LAYER_INFO_OVR_STYLE_COUNT               44
#define LAYER_INFO_OVR_LBL_COUNT                 45
#define LAYER_INFO_OVR_STYLE_CURRENT             46
#define LAYER_INFO_OVR_LBL_CURRENT               47
#define LAYER_INFO_OVR_LINE_COUNT                48
#define LAYER_INFO_OVR_PEN_COUNT                 49
#define LAYER_INFO_OVR_BRUSH_COUNT               50
#define LAYER_INFO_OVR_SYMBOL_COUNT              51
#define LAYER_INFO_OVR_FONT_COUNT                52
#define LAYER_INFO_TILE_SERVER_LEVEL             53
// Advanced labelling options
#define LAYER_INFO_LBL_AUTO_POSITION             54
#define LAYER_INFO_LBL_AUTO_SIZES                55
#define LAYER_INFO_LBL_SUPPRESS_IF_NO_FIT        56
#define LAYER_INFO_LBL_AUTO_SIZE_STEP            57
#define LAYER_INFO_LBL_CURVED_BEST_POSITION      58
#define LAYER_INFO_LBL_CURVED_FALLBACK           59
#define LAYER_INFO_LBL_USE_ABBREVIATION          60
#define LAYER_INFO_ABBREVIATION_EXPR             61
#define LAYER_INFO_LBL_AUTO_CALLOUT              62
#define LAYER_INFO_LBL_ORDER                     63
#define LAYER_INFO_FRIENDLYNAME                  64
#define LAYER_INFO_GRIDLINE_COORDSYS             65
#define LAYER_INFO_GRIDLINE_SPACE_HORIZONTAL     66
#define LAYER_INFO_GRIDLINE_SPACE_VERTICAL       67
#define LAYER_INFO_GRIDLINE_AUTO                 68
#define LAYER_INFO_GRIDLINE_MIN_HORIZONTAL       69
#define LAYER_INFO_GRIDLINE_MIN_VERTICAL         70
#define LAYER_INFO_GRIDLINE_PAPER_HORIZONTAL     71
#define LAYER_INFO_GRIDLINE_PAPER_VERTICAL       72
#define LAYER_INFO_GRIDLINE_OFFSET               73
#define LAYER_INFO_LABEL_POS_RETRY               74
#define LAYER_INFO_GRIDLINE_HORIZONTAL_LABEL_POS 75
#define LAYER_INFO_GRIDLINE_VERTICAL_LABEL_POS   76
#define LAYER_INFO_GRIDLINE_LABEL_DIRECTION      77
#define LAYER_INFO_ALIAS                         78
#define LAYER_INFO_LBL_LINE_PEN                  79
#define LAYER_INFO_RASTER_CONTRAST               80
#define LAYER_INFO_RASTER_BRIGHTNESS             81
#define LAYER_INFO_RASTER_ALPHA                  82
#define LAYER_INFO_RASTER_GREYSCALE              83
#define LAYER_INFO_LABEL_EXPR_TABLE              84
#define LAYER_INFO_LBL_CURVED_PERCENT_OVER       85
#define LAYER_INFO_LEGEND                        86

//----------------------------------------------------------------------------
// Values returned by LayerInfo() for LAYER_INFO_LABEL_ORIENTATION and
// LABEL_INFO_ORIENTATION.
//----------------------------------------------------------------------------
#define LAYER_INFO_LABEL_ORIENT_HORIZONTAL       0
#define LAYER_INFO_LABEL_ORIENT_PARALLEL         1
#define LAYER_INFO_LABEL_ORIENT_CURVED           2

//----------------------------------------------------------------------------
// Display Modes, returned by LayerInfo() for LAYER_INFO_DISPLAY
//----------------------------------------------------------------------------
#define LAYER_INFO_DISPLAY_OFF                   0
#define LAYER_INFO_DISPLAY_GRAPHIC               1
#define LAYER_INFO_DISPLAY_GLOBAL                2
#define LAYER_INFO_DISPLAY_VALUE                 3

//----------------------------------------------------------------------------
// Label Linetypes, returned by LayerInfo() for LAYER_INFO_LBL_LT
//----------------------------------------------------------------------------
#define LAYER_INFO_LBL_LT_NONE                   0
#define LAYER_INFO_LBL_LT_SIMPLE                 1
#define LAYER_INFO_LBL_LT_ARROW                  2

//----------------------------------------------------------------------------
// Label Positions, returned by LayerInfo() for LAYER_INFO_LBL_POS and
// LabelInfo() for LABEL_INFO_POSITION
//----------------------------------------------------------------------------
#define LAYER_INFO_LBL_POS_AUTO                 -1
#define LAYER_INFO_LBL_POS_CC                    0
#define LAYER_INFO_LBL_POS_TL                    1
#define LAYER_INFO_LBL_POS_TC                    2
#define LAYER_INFO_LBL_POS_TR                    3
#define LAYER_INFO_LBL_POS_CL                    4
#define LAYER_INFO_LBL_POS_CR                    5
#define LAYER_INFO_LBL_POS_BL                    6
#define LAYER_INFO_LBL_POS_BC                    7
#define LAYER_INFO_LBL_POS_BR                    8

//----------------------------------------------------------------------------
// Layer Types, returned by LayerInfo() for LAYER_INFO_TYPE
//----------------------------------------------------------------------------
#define LAYER_INFO_TYPE_NORMAL                   0
#define LAYER_INFO_TYPE_COSMETIC                 1
#define LAYER_INFO_TYPE_IMAGE                    2
#define LAYER_INFO_TYPE_THEMATIC                 3
#define LAYER_INFO_TYPE_GRID                     4
#define LAYER_INFO_TYPE_WMS                      5
#define LAYER_INFO_TYPE_TILESERVER               6
#define LAYER_INFO_TYPE_GRIDLINE                 7

//----------------------------------------------------------------------------
// Label visibility modes, from LayerInfo() for LAYER_INFO_LBL_VISIBILITY
//----------------------------------------------------------------------------
#define LAYER_INFO_LBL_VIS_OFF                   1
#define LAYER_INFO_LBL_VIS_ZOOM                  2
#define LAYER_INFO_LBL_VIS_ON                    3

//----------------------------------------------------------------------------
// Values Returned by LayerInfo() for LAYER_INFO_GRIDLINE_HORIZONTAL_LABEL_POS
// and LAYER_INFO_GRIDLINE_VERTICAL_LABEL_POS.
//-----------------------------------------------------------------------------
#define GRIDLINE_HORIZONTAL_LABEL_NONE  0
#define GRIDLINE_HORIZONTAL_LABEL_LEFT  1
#define GRIDLINE_HORIZONTAL_LABEL_RIGHT 2
#define GRIDLINE_HORIZONTAL_LABEL_BOTH  3

#define GRIDLINE_VERTICAL_LABEL_NONE    0
#define GRIDLINE_VERTICAL_LABEL_TOP     1
#define GRIDLINE_VERTICAL_LABEL_BOTTOM  2
#define GRIDLINE_VERTICAL_LABEL_BOTH    3

//----------------------------------------------------------------------------
// Values Returned by LayerInfo() for LAYER_INFO_GRIDLINE_LABEL_DIRECTION.
//----------------------------------------------------------------------------
#define GRIDLINE_LABEL_DIRECTION_PLUSMINUS 0
#define GRIDLINE_LABEL_DIRECTION_COMPASS   1

//----------------------------------------------------------------------------
// Code passed to StyleOverrideInfo function to get override style information
//----------------------------------------------------------------------------
#define STYLE_OVR_INFO_NAME                      1
#define STYLE_OVR_INFO_VISIBILITY                2
#define STYLE_OVR_INFO_ZOOM_MIN                  3
#define STYLE_OVR_INFO_ZOOM_MAX                  4
#define STYLE_OVR_INFO_ARROWS                    5
#define STYLE_OVR_INFO_NODES                     6
#define STYLE_OVR_INFO_CENTROIDS                 7
#define STYLE_OVR_INFO_ALPHA                     8
#define STYLE_OVR_INFO_TRANSLUCENCY              9
#define STYLE_OVR_INFO_LINE                      10
#define STYLE_OVR_INFO_PEN                       11
#define STYLE_OVR_INFO_BRUSH                     12
#define STYLE_OVR_INFO_SYMBOL                    13
#define STYLE_OVR_INFO_FONT                      14
#define STYLE_OVR_INFO_SYMBOL_COUNT              15
#define STYLE_OVR_INFO_LINE_COUNT                16
#define STYLE_OVR_INFO_PEN_COUNT                 17
#define STYLE_OVR_INFO_BRUSH_COUNT               18
#define STYLE_OVR_INFO_FONT_COUNT                19

//----------------------------------------------------------------------------
// Possible return value of StyleOverrideInfo for code STYLE_OVR_INFO_VISIBILITY
//----------------------------------------------------------------------------
#define STYLE_OVR_INFO_VIS_OFF                   0
#define STYLE_OVR_INFO_VIS_ON                    1
#define STYLE_OVR_INFO_VIS_OFF_ZOOM              2

//----------------------------------------------------------------------------
// Code passed to LabelOverrideInfo function to get override label information
//----------------------------------------------------------------------------
#define LBL_OVR_INFO_NAME                        1
#define LBL_OVR_INFO_VISIBILITY                  2
#define LBL_OVR_INFO_ZOOM_MIN                    3
#define LBL_OVR_INFO_ZOOM_MAX                    4
#define LBL_OVR_INFO_EXPR                        5
#define LBL_OVR_INFO_LT                          6
#define LBL_OVR_INFO_FONT                        7
#define LBL_OVR_INFO_PARALLEL                    8
#define LBL_OVR_INFO_POS                         9
#define LBL_OVR_INFO_OVERLAP                     10
#define LBL_OVR_INFO_DUPLICATES                  11
#define LBL_OVR_INFO_OFFSET                      12
#define LBL_OVR_INFO_MAX                         13
#define LBL_OVR_INFO_PARTIALSEGS                 14
#define LBL_OVR_INFO_ORIENTATION                 15
#define LBL_OVR_INFO_ALPHA                       16
#define LBL_OVR_INFO_AUTODISPLAY                 17
#define LBL_OVR_INFO_POS_RETRY                   18
#define LBL_OVR_INFO_LINE_PEN                    19
#define LBL_OVR_INFO_PERCENT_OVER                20
#define LBL_OVR_INFO_AUTO_POSITION               21
#define LBL_OVR_INFO_AUTO_SIZES                  22
#define LBL_OVR_INFO_SUPPRESS_IF_NO_FIT          23
#define LBL_OVR_INFO_AUTO_SIZE_STEP              24
#define LBL_OVR_INFO_CURVED_BEST_POSITION        25
#define LBL_OVR_INFO_CURVED_FALLBACK             26
#define LBL_OVR_INFO_USE_ABBREVIATION            27
#define LBL_OVR_INFO_ABBREVIATION_EXPR           28
#define LBL_OVR_INFO_AUTO_CALLOUT                29
#define LBL_OVR_INFO_EXPR_EDITED                 30
#define LBL_OVR_INFO_LT_EDITED                   31
#define LBL_OVR_INFO_FONT_EDITED                 32
#define LBL_OVR_INFO_PARALLEL_EDITED             33
#define LBL_OVR_INFO_POS_EDITED                  34
#define LBL_OVR_INFO_OVERLAP_EDITED              35
#define LBL_OVR_INFO_DUPLICATES_EDITED           36
#define LBL_OVR_INFO_OFFSET_EDITED               37
#define LBL_OVR_INFO_MAX_EDITED                  38
#define LBL_OVR_INFO_PARTIALSEGS_EDITED          39
#define LBL_OVR_INFO_ORIENTATION_EDITED          40
#define LBL_OVR_INFO_ALPHA_EDITED                41
#define LBL_OVR_INFO_AUTODISPLAY_EDITED          42
#define LBL_OVR_INFO_POS_RETRY_EDITED            43
#define LBL_OVR_INFO_LINE_PEN_EDITED             44
#define LBL_OVR_INFO_PERCENT_OVER_EDITED         45
#define LBL_OVR_INFO_AUTO_POSITION_EDITED        46
#define LBL_OVR_INFO_AUTO_SIZES_EDITED           47
#define LBL_OVR_INFO_SUPPRESS_IF_NO_FIT_EDITED   48
#define LBL_OVR_INFO_AUTO_SIZESTEP_EDITED        49
#define LBL_OVR_INFO_CURV_BEST_POS_EDITED        50
#define LBL_OVR_INFO_CURV_FALLBACK_EDITED        51
#define LBL_OVR_INFO_USE_ABBREV_EDITED           52
#define LBL_OVR_INFO_ABBREV_EXPR_EDITED          53
#define LBL_OVR_INFO_AUTO_CALLOUT_EDITED         54

//----------------------------------------------------------------------------
// Possible return value of LabelOverrideInfo for code LBL_OVR_INFO_VISIBILITY
//----------------------------------------------------------------------------
#define LBL_OVR_INFO_VIS_OFF                     0
#define LBL_OVR_INFO_VIS_ON                      1
#define LBL_OVR_INFO_VIS_OFF_ZOOM                2

//----------------------------------------------------------------------------
// LayerControlInfo() defines
//----------------------------------------------------------------------------
#define LC_INFO_SEL_COUNT                        1

//----------------------------------------------------------------------------
// LayerControlSelectionInfo() defines
//----------------------------------------------------------------------------
#define LC_SEL_INFO_NAME                         1
#define LC_SEL_INFO_TYPE                         2
#define LC_SEL_INFO_MAPWIN_ID                    3
#define LC_SEL_INFO_LAYER_ID                     4
#define LC_SEL_INFO_OVR_ID                       5
#define LC_SEL_INFO_HAS_COLOR_SAMPLE             6

//----------------------------------------------------------------------------
// Values returned by LayerControlSelectionInfo() for LC_SEL_INFO_TYPE.
//----------------------------------------------------------------------------
#define LC_SEL_INFO_TYPE_MAP                     0
#define LC_SEL_INFO_TYPE_LAYER                   1
#define LC_SEL_INFO_TYPE_GROUPLAYER              2
#define LC_SEL_INFO_TYPE_STYLE_OVR               3
#define LC_SEL_INFO_TYPE_LABEL_OVR               4
#define LC_SEL_INFO_TYPE_THEME_ITEM              5

//----------------------------------------------------------------------------
// HotlinkInfo() defines
//----------------------------------------------------------------------------
#define HOTLINK_INFO_EXPR                        1
#define HOTLINK_INFO_MODE                        2
#define HOTLINK_INFO_RELATIVE                    3
#define HOTLINK_INFO_ENABLED                     4
#define HOTLINK_INFO_ALIAS                       5

//----------------------------------------------------------------------------
// Hotlink activation modes, from LayerInfo() for LAYER_INFO_HOTLINK_MODE
//----------------------------------------------------------------------------
#define HOTLINK_MODE_LABEL                       0
#define HOTLINK_MODE_OBJ                         1
#define HOTLINK_MODE_BOTH                        2

//============================================================================
// LegendInfo() defines
//============================================================================
#define LEGEND_INFO_MAP_ID                       1
#define LEGEND_INFO_ORIENTATION                  2
#define LEGEND_INFO_NUM_FRAMES                   3
#define LEGEND_INFO_STYLE_SAMPLE_SIZE            4
#define LEGEND_INFO_LINE_SAMPLE_WIDTH            5
#define LEGEND_INFO_REGION_SAMPLE_WIDTH          6
#define LEGEND_INFO_REGION_SAMPLE_HEIGHT         7
#define LEGEND_INFO_DEFAULT_TITLE                9
#define LEGEND_INFO_DEFAULT_TITLE_FONT           10
#define LEGEND_INFO_DEFAULT_SUBTITLE             11
#define LEGEND_INFO_DEFAULT_SUBTITLE_FONT        12
#define LEGEND_INFO_DEFAULT_STYLE                13
#define LEGEND_INFO_DEFAULT_STYLE_FONT           14
#define LEGEND_INFO_DEFAULT_AUTO_FNT_SZ          15

//============================================================================
// Orientation codes, returned by LegendInfo() for LEGEND_INFO_ORIENTATION
//============================================================================
#define ORIENTATION_PORTRAIT                     1
#define ORIENTATION_LANDSCAPE                    2
#define ORIENTATION_CUSTOM                       3

//----------------------------------------------------------------------------
// Style sample codes, from LegendInfo() for LEGEND_INFO_STYLE_SAMPLE_SIZE
//----------------------------------------------------------------------------
#define STYLE_SAMPLE_SIZE_SMALL                  0
#define STYLE_SAMPLE_SIZE_LARGE                  1

//============================================================================
// LegendFrameInfo() defines
//============================================================================
#define FRAME_INFO_TYPE                          1
#define FRAME_INFO_MAP_LAYER_ID                  2
#define FRAME_INFO_REFRESHABLE                   3
#define FRAME_INFO_POS_X                         4
#define FRAME_INFO_POS_Y                         5
#define FRAME_INFO_WIDTH                         6
#define FRAME_INFO_HEIGHT                        7
#define FRAME_INFO_TITLE                         8
#define FRAME_INFO_TITLE_FONT                    9
#define FRAME_INFO_SUBTITLE                      10
#define FRAME_INFO_SUBTITLE_FONT                 11
#define FRAME_INFO_BORDER_PEN                    12
#define FRAME_INFO_NUM_STYLES                    13
#define FRAME_INFO_VISIBLE                       14
#define FRAME_INFO_COLUMN                        15
#define FRAME_INFO_LABEL                         16
#define FRAME_INFO_COLUMNS                       17
#define FRAME_INFO_NUM_VISIBLE_ROWS              18
#define FRAME_INFO_LINE_SAMPLE_WIDTH             19
#define FRAME_INFO_REGION_SAMPLE_WIDTH           20
#define FRAME_INFO_REGION_SAMPLE_HEIGHT          21
#define FRAME_INFO_AUTO_FONT_SIZE                22
#define FRAME_INFO_SHOW_COUNTS                   23
#define FRAME_INFO_DISPLAY_TYPE                  24
#define FRAME_INFO_BAR_THICKNESS                 25
#define FRAME_INFO_BAR_LENGTH                    26
#define FRAME_INFO_BAR_IS_HORIZONTAL             27

//============================================================================
// Frame Types, returned by LegendFrameInfo() for FRAME_INFO_TYPE
//============================================================================
#define FRAME_TYPE_NONE                          0
#define FRAME_TYPE_STYLE                         1
#define FRAME_TYPE_THEME                         2
#define FRAME_TYPE_RASTER                        3

//----------------------------------------------------------------------------
// Frame display type codes, from LegendFrameInfo() for FRAME_INFO_DISPLAY_TYPE
//----------------------------------------------------------------------------
#define FRAME_DISPLAY_TYPE_NOT_SPECIFIED         0
#define FRAME_DISPLAY_TYPE_GAP                   1
#define FRAME_DISPLAY_TYPE_PALETTE               2
#define FRAME_DISPLAY_TYPE_BAR                   3

//============================================================================
// Geocode Attributes, returned by GeocodeInfo()
//============================================================================
#define GEOCODE_STREET_NAME                      1
#define GEOCODE_STREET_NUMBER                    2
#define GEOCODE_MUNICIPALITY                     3
#define GEOCODE_MUNICIPALITY2                    4
#define GEOCODE_COUNTRY_SUBDIVISION              5
#define GEOCODE_COUNTRY_SUBDIVISION2             6
#define GEOCODE_POSTAL_CODE                      7

#define GEOCODE_DICTIONARY                       9
#define GEOCODE_BATCH_SIZE                       10
#define GEOCODE_FALLBACK_GEOGRAPHIC              11
#define GEOCODE_FALLBACK_POSTAL                  12
#define GEOCODE_OFFSET_CENTER                    13
#define GEOCODE_OFFSET_CENTER_UNITS              14
#define GEOCODE_OFFSET_END                       15
#define GEOCODE_OFFSET_END_UNITS                 16
#define GEOCODE_MIXED_CASE                       17
#define GEOCODE_RESULT_MARK_MULTIPLE             18
#define GEOCODE_COUNT_GEOCODED                   19
#define GEOCODE_COUNT_NOTGEOCODED                20
#define GEOCODE_UNABLE_TO_CONVERT_DATA           21
#define GEOCODE_MAX_BATCH_SIZE                   22

#define GEOCODE_PASSTHROUGH                      100

#define DICTIONARY_ALL                           1
#define DICTIONARY_ADDRESS_ONLY                  2
#define DICTIONARY_USER_ONLY                     3
#define DICTIONARY_PREFER_ADDRESS                4
#define DICTIONARY_PREFER_USER                   5

//============================================================================
// ISOGRAM Attributes, returned by IsogramInfo()
//============================================================================
#define ISOGRAM_BANDING                          1
#define ISOGRAM_MAJOR_ROADS_ONLY                 2
#define ISOGRAM_RETURN_HOLES                     3
#define ISOGRAM_MAJOR_POLYGON_ONLY               4
#define ISOGRAM_MAX_OFFROAD_DIST                 5
#define ISOGRAM_MAX_OFFROAD_DIST_UNITS           6
#define ISOGRAM_SIMPLIFICATION_FACTOR            7
#define ISOGRAM_DEFAULT_AMBIENT_SPEED            8
#define ISOGRAM_AMBIENT_SPEED_DIST_UNIT          9
#define ISOGRAM_AMBIENT_SPEED_TIME_UNIT          10
#define ISOGRAM_PROPAGATION_FACTOR               11
#define ISOGRAM_BATCH_SIZE                       12
#define ISOGRAM_POINTS_ONLY                      13
#define ISOGRAM_RECORDS_INSERTED                 14
#define ISOGRAM_RECORDS_NOTINSERTED              15
#define ISOGRAM_MAX_BATCH_SIZE                   16
#define ISOGRAM_MAX_BANDS                        17
#define ISOGRAM_MAX_DISTANCE                     18
#define ISOGRAM_MAX_DISTANCE_UNITS               19
#define ISOGRAM_MAX_TIME                         20
#define ISOGRAM_MAX_TIME_UNITS                   21

//============================================================================
// LegendStyleInfo() defines
//============================================================================
#define LEGEND_STYLE_INFO_TEXT                   1
#define LEGEND_STYLE_INFO_FONT                   2
#define LEGEND_STYLE_INFO_OBJ                    3
#define LEGEND_STYLE_INFO_ROW_VISIBLE            4
#define LEGEND_STYLE_INFO_COUNT                  5

//============================================================================
// LayoutInfo() defines
//============================================================================
#define LAYOUT_INFO_NUM_ITEMS                    1
#define LAYOUT_INFO_WIDTH                        2
#define LAYOUT_INFO_HEIGHT                       3
#define LAYOUT_INFO_LEFT_MARGIN                  4
#define LAYOUT_INFO_RIGHT_MARGIN                 5
#define LAYOUT_INFO_TOP_MARGIN                   6
#define LAYOUT_INFO_BOTTOM_MARGIN                7
#define LAYOUT_INFO_ZOOM                         8
#define LAYOUT_INFO_CENTER_X                     9
#define LAYOUT_INFO_CENTER_Y                     10
#define LAYOUT_INFO_NUM_PAGES                    11
#define LAYOUT_INFO_CUR_PAGE                     12
#define LAYOUT_INFO_GRID_VISIBILITY              13
#define LAYOUT_INFO_GRID_SNAP                    14
#define LAYOUT_INFO_GRID_SIZE                    15
#define LAYOUT_INFO_SMART_GUIDES                 16

//============================================================================
// LayoutItemInfo() and LayoutPageItemInfo() defines
//============================================================================
#define LAYOUT_ITEM_INFO_POS_X                             1
#define LAYOUT_ITEM_INFO_POS_Y                             2
#define LAYOUT_ITEM_INFO_WIDTH                             3
#define LAYOUT_ITEM_INFO_HEIGHT                            4
#define LAYOUT_ITEM_INFO_WIN                               5
#define LAYOUT_ITEM_INFO_SELECTED                          6
#define LAYOUT_ITEM_INFO_ACTIVATED                         7
#define LAYOUT_ITEM_INFO_EMPTY                             8
#define LAYOUT_ITEM_INFO_LEGEND_FRAME_ID                   9
#define LAYOUT_ITEM_INFO_LEGEND_DESIGNER_WINDOW            10
#define LAYOUT_ITEM_INFO_TYPE                              11
#define LAYOUT_ITEM_INFO_IMAGE_FILE                        12
#define LAYOUT_ITEM_INFO_OBJ                               13
#define LAYOUT_ITEM_INFO_NAME                              14
#define LAYOUT_ITEM_INFO_PEN                               15
#define LAYOUT_ITEM_INFO_FILL_BRUSH                        16
#define LAYOUT_ITEM_INFO_Z_ORDER                           17
#define LAYOUT_ITEM_INFO_ANGLE                             18
#define LAYOUT_ITEM_INFO_CUSTOM_FRAME_ADDIN_FILE_NAME      19
#define LAYOUT_ITEM_INFO_CUSTOM_FRAME_SERIALIZED_FILE_NAME 20
#define LAYOUT_ITEM_INFO_CUSTOM_FRAME_XPS_VIEW_FILE_NAME   21
#define LAYOUT_ITEM_INFO_NUM_ITEMS_ON_PAGE                 22
#define LAYOUT_ITEM_INFO_FRAME_UNIQUE_ID                   23
#define LAYOUT_ITEM_INFO_TEXTSTRING                        24
#define LAYOUT_ITEM_INFO_SMART_TEXTSTRING                  25

//============================================================================
// Item types, returned by LayoutItemInfo() and LayoutPageItemInfo() for
// LAYOUT_ITEM_INFO_TYPE
//============================================================================
#define LAYOUT_ITEM_TYPE_ANY                      -1
#define LAYOUT_ITEM_TYPE_EMPTY                    0
#define LAYOUT_ITEM_TYPE_MAPPER                   1
#define LAYOUT_ITEM_TYPE_BROWSER                  2
#define LAYOUT_ITEM_TYPE_LEGEND                   3
#define LAYOUT_ITEM_TYPE_TEXT                     4
#define LAYOUT_ITEM_TYPE_SHAPE                    5
#define LAYOUT_ITEM_TYPE_IMAGE                    6
#define LAYOUT_ITEM_TYPE_CUSTOM                   7

//============================================================================
// CurContext() defines
//============================================================================
#define CONTEXT_WINDOWID                          1
#define CONTEXT_PAGENUMBER                        2
#define CONTEXT_FRAMENUMBER                       3

//============================================================================
// LocateFile$() defines
//============================================================================

#define LOCATE_PREF_FILE                         0
#define LOCATE_DEF_WOR                           1
#define LOCATE_CLR_FILE                          2
#define LOCATE_PEN_FILE                          3
#define LOCATE_FNT_FILE                          4
#define LOCATE_ABB_FILE                          5
#define LOCATE_PRJ_FILE                          6
#define LOCATE_MNU_FILE                          7
#define LOCATE_CUSTSYMB_DIR                      8
#define LOCATE_THMTMPLT_DIR                      9
#define LOCATE_GRAPH_DIR                         10
#define LOCATE_WMS_SERVERLIST                    11
#define LOCATE_WFS_SERVERLIST                    12
#define LOCATE_GEOCODE_SERVERLIST                13
#define LOCATE_ROUTING_SERVERLIST                14
#define LOCATE_LAYOUT_TEMPLATE_DIR               15

//============================================================================
// GetLicenseKind() defines
//============================================================================
#define LICENSE_KIND_NOLICENSE                   0
#define LICENSE_KIND_REGULAR                     1
#define LICENSE_KIND_ADVANCE                     2

//============================================================================
// TableConnectionInfo() defines
//============================================================================
#define TABLE_CONNECTION_INFO_NUMBER             1

//============================================================================
// TableSchemaInfo() defines
//============================================================================
#define TABLE_SCHEMA_INFO_NAME                   1

//============================================================================
// Map3DInfo() defines
//============================================================================
#define MAP3D_INFO_SCALE                         1
#define MAP3D_INFO_RESOLUTION_X                  2
#define MAP3D_INFO_RESOLUTION_Y                  3
#define MAP3D_INFO_BACKGROUND                    4
#define MAP3D_INFO_UNITS                         5
#define MAP3D_INFO_LIGHT_X                       6
#define MAP3D_INFO_LIGHT_Y                       7
#define MAP3D_INFO_LIGHT_Z                       8
#define MAP3D_INFO_LIGHT_COLOR                   9
#define MAP3D_INFO_CAMERA_X                      10
#define MAP3D_INFO_CAMERA_Y                      11
#define MAP3D_INFO_CAMERA_Z                      12
#define MAP3D_INFO_CAMERA_FOCAL_X                13
#define MAP3D_INFO_CAMERA_FOCAL_Y                14
#define MAP3D_INFO_CAMERA_FOCAL_Z                15
#define MAP3D_INFO_CAMERA_VU_1                   16
#define MAP3D_INFO_CAMERA_VU_2                   17
#define MAP3D_INFO_CAMERA_VU_3                   18
#define MAP3D_INFO_CAMERA_VPN_1                  19
#define MAP3D_INFO_CAMERA_VPN_2                  20
#define MAP3D_INFO_CAMERA_VPN_3                  21
#define MAP3D_INFO_CAMERA_CLIP_NEAR              22
#define MAP3D_INFO_CAMERA_CLIP_FAR               23

//============================================================================
// MapperInfo() defines
//============================================================================
#define MAPPER_INFO_ZOOM                         1
#define MAPPER_INFO_SCALE                        2
#define MAPPER_INFO_CENTERX                      3
#define MAPPER_INFO_CENTERY                      4
#define MAPPER_INFO_MINX                         5
#define MAPPER_INFO_MINY                         6
#define MAPPER_INFO_MAXX                         7
#define MAPPER_INFO_MAXY                         8
#define MAPPER_INFO_LAYERS                       9
#define MAPPER_INFO_EDIT_LAYER                   10
#define MAPPER_INFO_XYUNITS                      11
#define MAPPER_INFO_DISTUNITS                    12
#define MAPPER_INFO_AREAUNITS                    13
#define MAPPER_INFO_SCROLLBARS                   14
#define MAPPER_INFO_DISPLAY                      15
#define MAPPER_INFO_NUM_THEMATIC                 16
#define MAPPER_INFO_COORDSYS_CLAUSE              17
#define MAPPER_INFO_COORDSYS_NAME                18
#define MAPPER_INFO_MOVE_DUPLICATE_NODES         19
#define MAPPER_INFO_DIST_CALC_TYPE               20
#define MAPPER_INFO_DISPLAY_DMS                  21
#define MAPPER_INFO_COORDSYS_CLAUSE_WITH_BOUNDS  22
#define MAPPER_INFO_CLIP_TYPE                    23
#define MAPPER_INFO_CLIP_REGION                  24
#define MAPPER_INFO_REPROJECTION                 25
#define MAPPER_INFO_RESAMPLING                   26
#define MAPPER_INFO_MERGE_MAP                    27
#define MAPPER_INFO_ALL_LAYERS                   28
#define MAPPER_INFO_GROUPLAYERS                  29
#define MAPPER_INFO_NUM_ADORNMENTS              200
#define MAPPER_INFO_ADORNMENT                   200
#define MAPPER_INFO_LABELS_SELECTABLE            30
#define MAPPER_INFO_HAS_CUSTOM_LABELS            31
#define MAPPER_INFO_CARTO_SCALE                  32
#define MAPPER_INFO_CARTO_SCALE_EXPR             33

//----------------------------------------------------------------------------
// Display Modes, returned by MapperInfo() for MAPPER_INFO_DISPLAY_DMS
//----------------------------------------------------------------------------
#define MAPPER_INFO_DISPLAY_DECIMAL              0
#define MAPPER_INFO_DISPLAY_DEGMINSEC            1
#define MAPPER_INFO_DISPLAY_MGRS                 2
#define MAPPER_INFO_DISPLAY_USNG_WGS84           3
#define MAPPER_INFO_DISPLAY_USNG_NAD27           4

//----------------------------------------------------------------------------
// Display Modes, returned by MapperInfo() for MAPPER_INFO_DISPLAY
//----------------------------------------------------------------------------
#define MAPPER_INFO_DISPLAY_SCALE                0
#define MAPPER_INFO_DISPLAY_ZOOM                 1
#define MAPPER_INFO_DISPLAY_POSITION             2
#define MAPPER_INFO_DISPLAY_CARTOGRAPHIC_SCALE   3

//----------------------------------------------------------------------------
// Distance Calculation Types from MapperInfo() for MAPPER_INFO_DIST_CALC_TYPE
//----------------------------------------------------------------------------
#define MAPPER_INFO_DIST_SPHERICAL               0
#define MAPPER_INFO_DIST_CARTESIAN               1

//----------------------------------------------------------------------------
// Clip Types, returned by MapperInfo() for MAPPER_INFO_CLIP_TYPE
//----------------------------------------------------------------------------
#define MAPPER_INFO_CLIP_DISPLAY_ALL             0
#define MAPPER_INFO_CLIP_DISPLAY_POLYOBJ         1
#define MAPPER_INFO_CLIP_OVERLAY                 2

//============================================================================
// MenuItemInfoByID() and MenuItemInfoByHandler() defines
//============================================================================
#define MENUITEM_INFO_ENABLED                    1
#define MENUITEM_INFO_CHECKED                    2
#define MENUITEM_INFO_CHECKABLE                  3
#define MENUITEM_INFO_SHOWHIDEABLE               4
#define MENUITEM_INFO_ACCELERATOR                5
#define MENUITEM_INFO_TEXT                       6
#define MENUITEM_INFO_HELPMSG                    7
#define MENUITEM_INFO_HANDLER                    8
#define MENUITEM_INFO_ID                         9

//============================================================================
// ObjectGeography() defines
//============================================================================
#define OBJ_GEO_MINX                             1
#define OBJ_GEO_LINEBEGX                         1
#define OBJ_GEO_POINTX                           1
#define OBJ_GEO_MINY                             2
#define OBJ_GEO_LINEBEGY                         2
#define OBJ_GEO_POINTY                           2
#define OBJ_GEO_MAXX                             3
#define OBJ_GEO_LINEENDX                         3
#define OBJ_GEO_MAXY                             4
#define OBJ_GEO_LINEENDY                         4
#define OBJ_GEO_ARCBEGANGLE                      5
#define OBJ_GEO_TEXTLINEX                        5
#define OBJ_GEO_ROUNDRADIUS                      5
#define OBJ_GEO_CENTROID                         5
#define OBJ_GEO_ARCENDANGLE                      6
#define OBJ_GEO_TEXTLINEY                        6
#define OBJ_GEO_TEXTANGLE                        7
#define OBJ_GEO_POINTZ                           8
#define OBJ_GEO_POINTM                           9

//============================================================================
// ObjectInfo() defines
//============================================================================
#define OBJ_INFO_TYPE                            1
#define OBJ_INFO_PEN                             2
#define OBJ_INFO_SYMBOL                          2
#define OBJ_INFO_TEXTFONT                        2
#define OBJ_INFO_BRUSH                           3
#define OBJ_INFO_NPNTS                           20
#define OBJ_INFO_TEXTSTRING                      3
#define OBJ_INFO_SMOOTH                          4
#define OBJ_INFO_FRAMEWIN                        4
#define OBJ_INFO_NPOLYGONS                       21
#define OBJ_INFO_TEXTSPACING                     4
#define OBJ_INFO_TEXTJUSTIFY                     5
#define OBJ_INFO_FRAMETITLE                      6
#define OBJ_INFO_TEXTARROW                       6
#define OBJ_INFO_FILLFRAME                       7
#define OBJ_INFO_REGION                          8
#define OBJ_INFO_PLINE                           9
#define OBJ_INFO_MPOINT                          10
#define OBJ_INFO_NONEMPTY                        11
#define OBJ_INFO_Z_UNIT_SET                      12
#define OBJ_INFO_Z_UNIT                          13
#define OBJ_INFO_HAS_Z                           14
#define OBJ_INFO_HAS_M                           15

//----------------------------------------------------------------------------
// Object types, returned by ObjectInfo() for OBJ_INFO_TYPE
//----------------------------------------------------------------------------
#define OBJ_TYPE_ARC                             1
#define OBJ_TYPE_ELLIPSE                         2
#define OBJ_TYPE_LINE                            3
#define OBJ_TYPE_PLINE                           4
#define OBJ_TYPE_POINT                           5
#define OBJ_TYPE_FRAME                           6
#define OBJ_TYPE_REGION                          7
#define OBJ_TYPE_RECT                            8
#define OBJ_TYPE_ROUNDRECT                       9
#define OBJ_TYPE_TEXT                            10
#define OBJ_TYPE_MPOINT                          11
#define OBJ_TYPE_COLLECTION                      12

//----------------------------------------------------------------------------
//  Constants passed to RegionInfo.
//----------------------------------------------------------------------------
#define REGION_INFO_IS_CLOCKWISE                 1

//============================================================================
// PrismMapInfo() defines
//============================================================================

#define PRISMMAP_INFO_SCALE                      1
#define PRISMMAP_INFO_BACKGROUND                 4
#define PRISMMAP_INFO_LIGHT_X                    6
#define PRISMMAP_INFO_LIGHT_Y                    7
#define PRISMMAP_INFO_LIGHT_Z                    8
#define PRISMMAP_INFO_LIGHT_COLOR                9
#define PRISMMAP_INFO_CAMERA_X                   10
#define PRISMMAP_INFO_CAMERA_Y                   11
#define PRISMMAP_INFO_CAMERA_Z                   12
#define PRISMMAP_INFO_CAMERA_FOCAL_X             13
#define PRISMMAP_INFO_CAMERA_FOCAL_Y             14
#define PRISMMAP_INFO_CAMERA_FOCAL_Z             15
#define PRISMMAP_INFO_CAMERA_VU_1                16
#define PRISMMAP_INFO_CAMERA_VU_2                17
#define PRISMMAP_INFO_CAMERA_VU_3                18
#define PRISMMAP_INFO_CAMERA_VPN_1               19
#define PRISMMAP_INFO_CAMERA_VPN_2               20
#define PRISMMAP_INFO_CAMERA_VPN_3               21
#define PRISMMAP_INFO_CAMERA_CLIP_NEAR           22
#define PRISMMAP_INFO_CAMERA_CLIP_FAR            23
#define PRISMMAP_INFO_INFOTIP_EXPR               24

//============================================================================
// SearchInfo() defines
//============================================================================
#define SEARCH_INFO_TABLE                        1
#define SEARCH_INFO_ROW                          2

//============================================================================
// SelectionInfo() defines
//============================================================================
#define SEL_INFO_TABLENAME                       1
#define SEL_INFO_SELNAME                         2
#define SEL_INFO_NROWS                           3

//============================================================================
// Server statement and function defines
//============================================================================
//----------------------------------------------------------------------------
// Return Codes
//----------------------------------------------------------------------------
#define SRV_SUCCESS                              0
#define SRV_SUCCESS_WITH_INFO                    1
#define SRV_ERROR                                -1
#define SRV_INVALID_HANDLE                       -2
#define SRV_NEED_DATA                            99
#define SRV_NO_MORE_DATA                         100

//----------------------------------------------------------------------------
// Special values for the status associated with a fetched value
//----------------------------------------------------------------------------
#define SRV_NULL_DATA                            -1
#define SRV_TRUNCATED_DATA                       -2

//----------------------------------------------------------------------------
// Server_ColumnInfo() defines
//----------------------------------------------------------------------------
#define SRV_COL_INFO_NAME                        1
#define SRV_COL_INFO_TYPE                        2
#define SRV_COL_INFO_WIDTH                       3
#define SRV_COL_INFO_PRECISION                   4
#define SRV_COL_INFO_SCALE                       5
#define SRV_COL_INFO_VALUE                       6
#define SRV_COL_INFO_STATUS                      7
#define SRV_COL_INFO_ALIAS                       8

//----------------------------------------------------------------------------
// Column types, returned by Server_ColumnInfo() for SRV_COL_INFO_TYPE
//----------------------------------------------------------------------------
#define SRV_COL_TYPE_NONE                        0
#define SRV_COL_TYPE_CHAR                        1
#define SRV_COL_TYPE_DECIMAL                     2
#define SRV_COL_TYPE_INTEGER                     3
#define SRV_COL_TYPE_SMALLINT                    4
#define SRV_COL_TYPE_DATE                        5
#define SRV_COL_TYPE_LOGICAL                     6
#define SRV_COL_TYPE_FLOAT                       8
#define SRV_COL_TYPE_FIXED_LEN_STRING            16
#define SRV_COL_TYPE_BIN_STRING                  17

//----------------------------------------------------------------------------
// Server_DriverInfo() Attr defines
//----------------------------------------------------------------------------
#define SRV_DRV_INFO_NAME                        1
#define SRV_DRV_INFO_NAME_LIST                   2
#define SRV_DRV_DATA_SOURCE                      3

//----------------------------------------------------------------------------
// Server_ConnectInfo() Attr defines
//----------------------------------------------------------------------------
#define SRV_CONNECT_INFO_DRIVER_NAME             1
#define SRV_CONNECT_INFO_DB_NAME                 2
#define SRV_CONNECT_INFO_SQL_USER_ID             3
#define SRV_CONNECT_INFO_DS_NAME                 4
#define SRV_CONNECT_INFO_QUOTE_CHAR              5
#define SRV_CONNECT_INFO_DATABASE_NAME           6
#define SRV_CONNECT_INFO_SERVER_NAME             7

//----------------------------------------------------------------------------
// Fetch Directions (used by ServerFetch function in some code libraries)
//----------------------------------------------------------------------------
#define SRV_FETCH_NEXT                           -1
#define SRV_FETCH_PREV                           -2
#define SRV_FETCH_FIRST                          -3
#define SRV_FETCH_LAST                           -4

//----------------------------------------------------------------------------
//Oracle workspace manager
//----------------------------------------------------------------------------
#define SRV_WM_HIST_NONE                          0
#define SRV_WM_HIST_OVERWRITE                     1
#define SRV_WM_HIST_NO_OVERWRITE                  2

//============================================================================
// SessionInfo() defines
//============================================================================

#define SESSION_INFO_COORDSYS_CLAUSE             1
#define SESSION_INFO_DISTANCE_UNITS              2
#define SESSION_INFO_AREA_UNITS                  3
#define SESSION_INFO_PAPER_UNITS                 4
#define SESSION_INFO_DEF_NATIVE_BLKSZ            5
#define SESSION_INFO_DEF_NATIVEX_BLKSZ           6
#define SESSION_INFO_TEMP_PATH                   7

//============================================================================
// Set Next Document Style defines
//============================================================================
#define WIN_STYLE_STANDARD                       0
#define WIN_STYLE_CHILD                          1
#define WIN_STYLE_POPUP_FULLCAPTION              2
#define WIN_STYLE_POPUP                          3

//============================================================================
// StringCompare() defines
//============================================================================
#define STR_LT                                   -1
#define STR_GT                                   1
#define STR_EQ                                   0

//============================================================================
// StyleAttr() defines
//============================================================================
#define PEN_WIDTH                                1
#define PEN_PATTERN                              2
#define PEN_COLOR                                4
#define PEN_INDEX                                5
#define PEN_INTERLEAVED                          6
#define BRUSH_PATTERN                            1
#define BRUSH_FORECOLOR                          2
#define BRUSH_BACKCOLOR                          3
#define FONT_NAME                                1
#define FONT_STYLE                               2
#define FONT_POINTSIZE                           3
#define FONT_FORECOLOR                           4
#define FONT_BACKCOLOR                           5
#define SYMBOL_CODE                              1
#define SYMBOL_COLOR                             2
#define SYMBOL_POINTSIZE                         3
#define SYMBOL_ANGLE                             4
#define SYMBOL_FONT_NAME                         5
#define SYMBOL_FONT_STYLE                        6
#define SYMBOL_KIND                              7
#define SYMBOL_CUSTOM_NAME                       8
#define SYMBOL_CUSTOM_STYLE                      9

//----------------------------------------------------------------------------
// Symbol kinds returned by StyleAttr() for SYMBOL_KIND
//----------------------------------------------------------------------------
#define SYMBOL_KIND_VECTOR                       1
#define SYMBOL_KIND_FONT                         2
#define SYMBOL_KIND_CUSTOM                       3

//============================================================================
// SystemInfo() defines
//============================================================================
#define SYS_INFO_PLATFORM                        1
#define SYS_INFO_APPVERSION                      2
#define SYS_INFO_MIVERSION                       3
#define SYS_INFO_RUNTIME                         4
#define SYS_INFO_CHARSET                         5
#define SYS_INFO_COPYPROTECTED                   6
#define SYS_INFO_APPLICATIONWND                  7
#define SYS_INFO_DDESTATUS                       8
#define SYS_INFO_MAPINFOWND                      9
#define SYS_INFO_NUMBER_FORMAT                   10
#define SYS_INFO_DATE_FORMAT                     11
#define SYS_INFO_DIG_INSTALLED                   12
#define SYS_INFO_DIG_MODE                        13
#define SYS_INFO_MIPLATFORM                      14
#define SYS_INFO_MDICLIENTWND                    15
#define SYS_INFO_PRODUCTLEVEL                    16
#define SYS_INFO_APPIDISPATCH                    17
#define SYS_INFO_MIBUILD_NUM                     18
#define SYS_INFO_MIFULLVERSION                   19
#define SYS_INFO_IMAPINFOAPPLICATION             20
#define SYS_INFO_MAPINFO_INTERFACE               21
#define SYS_INFO_UNICODE                         22
#define SYS_INFO_CULTURE                         23
#define SYS_INFO_LOCALE                          24
#define SYS_INFO_MIRIBBONSTATE_VERSION           25
#define SYS_INFO_LICENSE_MODE                    26
#define SYS_INFO_SUBSCRIPTION_PLAN               27

//----------------------------------------------------------------------------
// Platform, returned by SystemInfo() for SYS_INFO_PLATFORM
//----------------------------------------------------------------------------
#define PLATFORM_SPECIAL                         0
#define PLATFORM_WIN                             1
#define PLATFORM_MAC                             2
#define PLATFORM_MOTIF                           3
#define PLATFORM_X11                             4
#define PLATFORM_XOL                             5

//----------------------------------------------------------------------------
// Version, returned by SystemInfo() for SYS_INFO_MIPLATFORM
//----------------------------------------------------------------------------
#define MIPLATFORM_SPECIAL                       0
#define MIPLATFORM_WIN16                         1
#define MIPLATFORM_WIN32                         2
#define MIPLATFORM_POWERMAC                      3
#define MIPLATFORM_MAC68K                        4
#define MIPLATFORM_HP                            5
#define MIPLATFORM_SUN                           6
#define MIPLATFORM_WIN64                         7

/****************************************************************************/
/* Codes returned by the SystemInfo function to describe the interface type  */
/* that the application is running on for SYS_INFO_MAPINFO_INTERFACE         */
/****************************************************************************/
#define MIINTERFACE_CLASSICMENU                  0
#define MIINTERFACE_RIBBON                       1

//============================================================================
// TableInfo() defines
//============================================================================
#define TAB_INFO_NAME                            1
#define TAB_INFO_NUM                             2
#define TAB_INFO_TYPE                            3
#define TAB_INFO_NCOLS                           4
#define TAB_INFO_MAPPABLE                        5
#define TAB_INFO_READONLY                        6
#define TAB_INFO_TEMP                            7
#define TAB_INFO_NROWS                           8
#define TAB_INFO_EDITED                          9
#define TAB_INFO_FASTEDIT                        10
#define TAB_INFO_UNDO                            11
#define TAB_INFO_MAPPABLE_TABLE                  12
#define TAB_INFO_USERMAP                         13
#define TAB_INFO_USERBROWSE                      14
#define TAB_INFO_USERCLOSE                       15
#define TAB_INFO_USEREDITABLE                    16
#define TAB_INFO_USERREMOVEMAP                   17
#define TAB_INFO_USERDISPLAYMAP                  18
#define TAB_INFO_TABFILE                         19
#define TAB_INFO_MINX                            20
#define TAB_INFO_MINY                            21
#define TAB_INFO_MAXX                            22
#define TAB_INFO_MAXY                            23
#define TAB_INFO_SEAMLESS                        24
#define TAB_INFO_COORDSYS_MINX                   25
#define TAB_INFO_COORDSYS_MINY                   26
#define TAB_INFO_COORDSYS_MAXX                   27
#define TAB_INFO_COORDSYS_MAXY                   28
#define TAB_INFO_COORDSYS_CLAUSE                 29
#define TAB_INFO_COORDSYS_NAME                   30
#define TAB_INFO_NREFS                           31
#define TAB_INFO_SUPPORT_MZ                      32
#define TAB_INFO_Z_UNIT_SET                      33
#define TAB_INFO_Z_UNIT                          34
#define TAB_INFO_BROWSER_LIST                    35
#define TAB_INFO_THEME_METADATA                  36
#define TAB_INFO_COORDSYS_CLAUSE_WITHOUT_BOUNDS  37
#define TAB_INFO_DESCRIPTION                     38
#define TAB_INFO_ID                              39
#define TAB_INFO_PARENTID                        40
#define TAB_INFO_ISMANAGED                       41
#define TAB_INFO_ADSK_TEXTOBJECT                 42
#define TAB_INFO_OVERRIDE_COORDINATE_ORDER       43
#define TAB_INFO_PERSIST                         44
#define TAB_INFO_PREFER_HTML_FOR_INFO_TOOL       45
#define TAB_INFO_INUSE                           46
#define TAB_INFO_DATA_FORMAT                     47
#define TAB_INFO_CHARSET                         48
#define TAB_INFO_MAP_BLOCKSIZE                   49
#define TAB_INFO_READVERSION                     50
#define TAB_INFO_EDITVERSION                     51
#define TAB_INFO_READALLVERSION                  52
#define TAB_INFO_LOCALE                          53
#define TAB_INFO_CONNECTIONS_COUNT               54
#define TAB_INFO_SCHEMAS_COUNT                   55
#define TAB_INFO_QUERY_STRING                    56
#define TAB_INFO_ISHIDDEN                        57
#define TAB_INFO_NUM_BASETABLES                  58
#define TAB_INFO_BASETABLE_ALIASES               59

//----------------------------------------------------------------------------
// Table type defines, returned by TableInfo() for TAB_INFO_TYPE
//----------------------------------------------------------------------------
#define TAB_TYPE_BASE                            1
#define TAB_TYPE_RESULT                          2
#define TAB_TYPE_VIEW                            3
#define TAB_TYPE_IMAGE                           4
#define TAB_TYPE_LINKED                          5
#define TAB_TYPE_WMS                             6
#define TAB_TYPE_WFS                             7
#define TAB_TYPE_FME                             8
#define TAB_TYPE_TILESERVER                      9

//----------------------------------------------------------------------------
// Table format defines, returned by TableInfo() for TAB_INFO_DATA_FORMAT
//----------------------------------------------------------------------------
#define TAB_DATA_FORMAT_NONE                     0
#define TAB_DATA_FORMAT_DBASE                    1
#define TAB_DATA_FORMAT_LOTUS                    2
#define TAB_DATA_FORMAT_ASCII                    3
#define TAB_DATA_FORMAT_EXCEL                    4
#define TAB_DATA_FORMAT_NATIVE                   5
#define TAB_DATA_FORMAT_RASTER                   6
#define TAB_DATA_FORMAT_DAO                      8
#define TAB_DATA_FORMAT_ODBC                     10
#define TAB_DATA_FORMAT_NATIVEX                  14
#define TAB_DATA_FORMAT_GPKG                     15
#define TAB_DATA_FORMAT_SHAPE                    16

//----------------------------------------------------------------------------
// ThemeInfo() defines
//----------------------------------------------------------------------------
#define THEME_INFO_TYPE                          1
#define THEME_INFO_ITEM_COUNT                    2
#define THEME_INFO_TABLE_ALIAS                   3
#define THEME_INFO_EXPRESSIONS_COUNT             4

//----------------------------------------------------------------------------
// Theme type defines returned by ThemeInfo() for THEME_INFO_TYPE
//----------------------------------------------------------------------------
#define THEME_TYPE_RANGES                        0
#define THEME_TYPE_BARS                          1
#define THEME_TYPE_PIES                          2
#define THEME_TYPE_GRADUATED                     3
#define THEME_TYPE_DOTDENSITY                    4
#define THEME_TYPE_UNIQUE                        5
#define THEME_TYPE_GRID                          6

//----------------------------------------------------------------------------
// ThemeItemInfo() defines
//----------------------------------------------------------------------------
#define THEME_ITEM_INFO_VISIBLE                  1
#define THEME_ITEM_INFO_LABEL                    2

//----------------------------------------------------------------------------
// ThemeExpressionInfo() defines
//----------------------------------------------------------------------------
#define THEME_EXPRESSION_INFO_STRING             1
#define THEME_EXPRESSION_INFO_TYPE               2

//----------------------------------------------------------------------------
// Column type defines, returned by ThemeExpressionInfo() for
// THEME_EXPRESSION_INFO_TYPE
//----------------------------------------------------------------------------
#define EXPRESSION_TYPE_CHAR                     1
#define EXPRESSION_TYPE_DECIMAL                  2
#define EXPRESSION_TYPE_INTEGER                  3
#define EXPRESSION_TYPE_SMALLINT                 4
#define EXPRESSION_TYPE_DATE                     5
#define EXPRESSION_TYPE_LOGICAL                  6
#define EXPRESSION_TYPE_GRAPHIC                  7
#define EXPRESSION_TYPE_FLOAT                    8
#define EXPRESSION_TYPE_TIME                     37
#define EXPRESSION_TYPE_DATETIME                 38

//----------------------------------------------------------------------------
// RangeThemeInfo() defines
//----------------------------------------------------------------------------
#define RANGETHEME_INFO_ITEMS_COUNT              1
#define RANGETHEME_INFO_VARY_STYLE               2
#define RANGETHEME_INFO_REPLACE_STYLE            3
#define RANGETHEME_INFO_OTHERS_LINE              4
#define RANGETHEME_INFO_OTHERS_PEN               5
#define RANGETHEME_INFO_OTHERS_BRUSH             6
#define RANGETHEME_INFO_OTHERS_SYMBOL            7
#define RANGETHEME_INFO_BASE_LINE                8
#define RANGETHEME_INFO_BASE_BRUSH               9
#define RANGETHEME_INFO_BASE_SYMBOL              10
#define RANGETHEME_INFO_METHOD                   11
#define RANGETHEME_INFO_ROUNDBY                  12

//----------------------------------------------------------------------------
// Range theme vary style by defines returned by RangeThemeInfo() for
// RANGETHEME_INFO_VARY_STYLE
//----------------------------------------------------------------------------
#define RANGETHEME_INFO_VARY_BY_ALL              0
#define RANGETHEME_INFO_VARY_BY_COLOR            1
#define RANGETHEME_INFO_VARY_BY_SIZE             2

//----------------------------------------------------------------------------
// RangeThemeItemInfo() defines
//----------------------------------------------------------------------------
#define RANGETHEME_ITEM_INFO_LINE                1
#define RANGETHEME_ITEM_INFO_PEN                 2
#define RANGETHEME_ITEM_INFO_BRUSH               3
#define RANGETHEME_ITEM_INFO_SYMBOL              4
#define RANGETHEME_ITEM_INFO_FROM                5
#define RANGETHEME_ITEM_INFO_TO                  6

//----------------------------------------------------------------------------
// IndividualThemeInfo() defines
//----------------------------------------------------------------------------
#define INDIVTHEME_INFO_ITEMS_COUNT              1
#define INDIVTHEME_INFO_VARY_STYLE               2
#define INDIVTHEME_INFO_REPLACE_STYLE            3
#define INDIVTHEME_INFO_OTHERS_LINE              4
#define INDIVTHEME_INFO_OTHERS_PEN               5
#define INDIVTHEME_INFO_OTHERS_BRUSH             6
#define INDIVTHEME_INFO_OTHERS_SYMBOL            7

//----------------------------------------------------------------------------
// Individual theme vary style by defines returned by IndividualThemeInfo()
// for INDIVTHEME_INFO_VARY_STYLE
//----------------------------------------------------------------------------
#define INDIVTHEME_INFO_VARY_BY_ALL              0
#define INDIVTHEME_INFO_VARY_BY_COLOR            1
#define INDIVTHEME_INFO_VARY_BY_SIZE             2

//----------------------------------------------------------------------------
// IndividualThemeItemInfo() defines
//----------------------------------------------------------------------------
#define INDIVTHEME_ITEM_INFO_LINE                1
#define INDIVTHEME_ITEM_INFO_PEN                 2
#define INDIVTHEME_ITEM_INFO_BRUSH               3
#define INDIVTHEME_ITEM_INFO_SYMBOL              4
#define INDIVTHEME_ITEM_INFO_VALUE               5

//----------------------------------------------------------------------------
// DotDensityThemeInfo() defines
//----------------------------------------------------------------------------
#define DOTDENSITYTHEME_INFO_CONCE               1
#define DOTDENSITYTHEME_INFO_SQUARE              2
#define DOTDENSITYTHEME_INFO_SIZE                3
#define DOTDENSITYTHEME_INFO_COLOR               4

//----------------------------------------------------------------------------
// GraduatedThemeInfo() defines
//----------------------------------------------------------------------------
#define GRADUATEDTHEME_INFO_SYMBOL               1
#define GRADUATEDTHEME_INFO_VALUE                2
#define GRADUATEDTHEME_INFO_SHOW_NEG             3
#define GRADUATEDTHEME_INFO_SHOW_NEG_SY          4
#define GRADUATEDTHEME_INFO_GRADUATION           5

//----------------------------------------------------------------------------
// Graduated theme graduation defines returned by GraduatedThemeInfo() for
// GRADUATEDTHEME_INFO_GRADUATION
//----------------------------------------------------------------------------
#define GRADUATEDTHEME_INFO_GRAD_ROOT            0
#define GRADUATEDTHEME_INFO_GRAD_CONST           1
#define GRADUATEDTHEME_INFO_GRAD_LOG             2

//----------------------------------------------------------------------------
// BarChartThemeInfo() defines
//----------------------------------------------------------------------------
#define BARCHARTTHEME_INFO_GRADUATION            1
#define BARCHARTTHEME_INFO_ORIENTATION           2
#define BARCHARTTHEME_INFO_PEN                   3
#define BARCHARTTHEME_INFO_CHART_TYPE            4
#define BARCHARTTHEME_INFO_GRAD_STACKS           5
#define BARCHARTTHEME_INFO_INDEP_SCALES          6
#define BARCHARTTHEME_INFO_HEIGHT                7
#define BARCHARTTHEME_INFO_HEIGHT_UNIT           8
#define BARCHARTTHEME_INFO_HEIGHT_VALUE          9
#define BARCHARTTHEME_INFO_WIDTH                 10
#define BARCHARTTHEME_INFO_WIDTH_UNIT            11
#define BARCHARTTHEME_INFO_BRUSH                 12

//----------------------------------------------------------------------------
// Bar chart theme graduation defines returned by BarChartThemeInfo() for
// BARCHARTTHEME_INFO_GRADUATION
//----------------------------------------------------------------------------
#define BARCHARTTHEME_INFO_GRAD_ROOT             0
#define BARCHARTTHEME_INFO_GRAD_CONST            1
#define BARCHARTTHEME_INFO_GRAD_LOG              2

//----------------------------------------------------------------------------
// Bar chart theme orientation defines returned by BarChartThemeInfo() for
// BARCHARTTHEME_INFO_ORIENTATION
//----------------------------------------------------------------------------
#define BARCHARTTHEME_INFO_ORI_CC                0
#define BARCHARTTHEME_INFO_ORI_TL                1
#define BARCHARTTHEME_INFO_ORI_TC                2
#define BARCHARTTHEME_INFO_ORI_TR                3
#define BARCHARTTHEME_INFO_ORI_CL                4
#define BARCHARTTHEME_INFO_ORI_CR                5
#define BARCHARTTHEME_INFO_ORI_BL                6
#define BARCHARTTHEME_INFO_ORI_BC                7
#define BARCHARTTHEME_INFO_ORI_BR                8

//----------------------------------------------------------------------------
// Bar chart theme chart type  defines returned by BarChartThemeInfo() for
// BARCHARTTHEME_INFO_CHART_TYPE
//----------------------------------------------------------------------------
#define BARCHARTTHEME_INFO_CHART_STA             0
#define BARCHARTTHEME_INFO_CHART_MUL             1

//----------------------------------------------------------------------------
// BarChartThemeItemInfo() defines
//----------------------------------------------------------------------------
#define BARCHARTTHEME_ITEM_INFO_BRUSH            1

//----------------------------------------------------------------------------
// PieChartThemeInfo() defines
//----------------------------------------------------------------------------
#define PIECHARTTHEME_INFO_GRADUATION            1
#define PIECHARTTHEME_INFO_ORIENTATION           2
#define PIECHARTTHEME_INFO_PEN                   3
#define PIECHARTTHEME_GRADUATED                  4
#define PIECHARTTHEME_HALF_PIE                   5
#define PIECHARTTHEME_DIAMETER                   6
#define PIECHARTTHEME_DIAMETER_UNIT              7
#define PIECHARTTHEME_DIAMETER_VALUE             8
#define PIECHARTTHEME_START_ANGLE                9
#define PIECHARTTHEME_CLOCKWISE                  10

//----------------------------------------------------------------------------
// Pie chart theme graduation defines returned by PieChartThemeInfo() for
// PIECHARTTHEME_INFO_GRADUATION
//----------------------------------------------------------------------------
#define PIECHARTTHEME_INFO_GRAD_ROOT             0
#define PIECHARTTHEME_INFO_GRAD_CONST            1
#define PIECHARTTHEME_INFO_GRAD_LOG              2

//----------------------------------------------------------------------------
// Pie chart theme orientation defines returned by PieChartThemeInfo() for
// PIECHARTTHEME_INFO_ORIENTATION
//----------------------------------------------------------------------------
#define PIECHARTTHEME_INFO_ORI_CC                0
#define PIECHARTTHEME_INFO_ORI_TL                1
#define PIECHARTTHEME_INFO_ORI_TC                2
#define PIECHARTTHEME_INFO_ORI_TR                3
#define PIECHARTTHEME_INFO_ORI_CL                4
#define PIECHARTTHEME_INFO_ORI_CR                5
#define PIECHARTTHEME_INFO_ORI_BL                6
#define PIECHARTTHEME_INFO_ORI_BC                7
#define PIECHARTTHEME_INFO_ORI_BR                8

//----------------------------------------------------------------------------
// PieChartThemeItemInfo() defines
//----------------------------------------------------------------------------
#define PIECHARTTHEME_ITEM_INFO_BRUSH            1

//****************************************************************************
// Defines used in LibraryServiceInfo function for what information to return.
//****************************************************************************
#define LIBSRVC_INFO_LIBSRVCMODE                 1
#define LIBSRVC_INFO_LIBVERSION                  2
#define LIBSRVC_INFO_DEFURLPATH                  3
#define LIBSRVC_INFO_LISTCSWURL                  4

//----------------------------------------------------------------------------
// TableListInfo() defines
//----------------------------------------------------------------------------
#define TL_INFO_SEL_COUNT                        1

//----------------------------------------------------------------------------
// TableListSelectionInfo() defines
//----------------------------------------------------------------------------
#define TL_SEL_INFO_NAME                         1
#define TL_SEL_INFO_ID                           2

//----------------------------------------------------------------------------
// RasterTableInfo() defines
//----------------------------------------------------------------------------
#define RASTER_TAB_INFO_IMAGE_NAME               1
#define RASTER_TAB_INFO_WIDTH                    2
#define RASTER_TAB_INFO_HEIGHT                   3
#define RASTER_TAB_INFO_IMAGE_TYPE               4
#define RASTER_TAB_INFO_BITS_PER_PIXEL           5
#define RASTER_TAB_INFO_IMAGE_CLASS              6
#define RASTER_TAB_INFO_NUM_CONTROL_POINTS       7
#define RASTER_TAB_INFO_BRIGHTNESS               8
#define RASTER_TAB_INFO_CONTRAST                 9
#define RASTER_TAB_INFO_GREYSCALE               10
#define RASTER_TAB_INFO_DISPLAY_TRANSPARENT     11
#define RASTER_TAB_INFO_TRANSPARENT_COLOR       12
#define RASTER_TAB_INFO_ALPHA                   13
#define RASTER_TAB_INFO_HANDLERNAME             14
#define RASTER_TAB_INFO_LEGEND                  15
#define RASTER_TAB_INFO_LEGEND_XML              16
#define RASTER_TAB_INFO_HANDLERFULLPATH         17
#define RASTER_TAB_INFO_INFLECTIONS             18

//----------------------------------------------------------------------------
// Image type defines returned by RasterTableInfo() for RASTER_TAB_INFO_IMAGE_TYPE
//----------------------------------------------------------------------------
#define IMAGE_TYPE_RASTER                        0
#define IMAGE_TYPE_GRID                          1

//----------------------------------------------------------------------------
// Image class defines returned by RasterTableInfo() for RASTER_TAB_INFO_IMAGE_CLASS
//----------------------------------------------------------------------------
#define IMAGE_CLASS_BILEVEL                      0
#define IMAGE_CLASS_GREYSCALE                    1
#define IMAGE_CLASS_PALETTE                      2
#define IMAGE_CLASS_RGB                          3

//----------------------------------------------------------------------------
//  GridTableInfo() defines
//----------------------------------------------------------------------------
#define GRID_TAB_INFO_MIN_VALUE                  1
#define GRID_TAB_INFO_MAX_VALUE                  2
#define GRID_TAB_INFO_HAS_HILLSHADE              3

//----------------------------------------------------------------------------
//  ControlPointInfo() defines
//----------------------------------------------------------------------------
#define RASTER_CONTROL_POINT_X                   1
#define RASTER_CONTROL_POINT_Y                   2
#define GEO_CONTROL_POINT_X                      3
#define GEO_CONTROL_POINT_Y                      4
#define TAB_GEO_CONTROL_POINT_X                  5
#define TAB_GEO_CONTROL_POINT_Y                  6

//============================================================================
// UndoInfo() defines
//============================================================================
#define UNDO_INFO_COUNT                          1
#define UNDO_INFO_UNDO_ACTION_NAME               2
#define UNDO_INFO_REDO_ACTION_NAME               3
#define UNDO_INFO_CAN_UNDO                       4
#define UNDO_INFO_CAN_REDO                       5
#define UNDO_INFO_ENABLED                        6
#define UNDO_INFO_CURRENT_POS                    7

//============================================================================
// WindowInfo() defines
//============================================================================
#define WIN_INFO_NAME                            1
#define WIN_INFO_TYPE                            3
#define WIN_INFO_WIDTH                           4
#define WIN_INFO_HEIGHT                          5
#define WIN_INFO_X                               6
#define WIN_INFO_Y                               7
#define WIN_INFO_TOPMOST                         8
#define WIN_INFO_STATE                           9
#define WIN_INFO_TABLE                           10
#define WIN_INFO_LEGENDS_MAP                     10
#define WIN_INFO_ADORNMENTS_MAP                  10
#define WIN_INFO_OPEN                            11
#define WIN_INFO_WND                             12
#define WIN_INFO_WINDOWID                        13
#define WIN_INFO_WORKSPACE                       14
#define WIN_INFO_CLONEWINDOW                     15
#define WIN_INFO_SYSMENUCLOSE                    16
#define WIN_INFO_AUTOSCROLL                      17
#define WIN_INFO_SMARTPAN                        18
#define WIN_INFO_SNAPMODE                        19
#define WIN_INFO_SNAPTHRESHOLD                   20
#define WIN_INFO_PRINTER_NAME                    21
#define WIN_INFO_PRINTER_ORIENT                  22
#define WIN_INFO_PRINTER_COPIES                  23
#define WIN_INFO_PRINTER_PAPERSIZE               24
#define WIN_INFO_PRINTER_LEFTMARGIN              25
#define WIN_INFO_PRINTER_RIGHTMARGIN             26
#define WIN_INFO_PRINTER_TOPMARGIN               27
#define WIN_INFO_PRINTER_BOTTOMMARGIN            28
#define WIN_INFO_PRINTER_BORDER                  29
#define WIN_INFO_PRINTER_TRUECOLOR               30
#define WIN_INFO_PRINTER_DITHER                  31
#define WIN_INFO_PRINTER_METHOD                  32
#define WIN_INFO_PRINTER_TRANSPRASTER            33
#define WIN_INFO_PRINTER_TRANSPVECTOR            34
#define WIN_INFO_EXPORT_BORDER                   35
#define WIN_INFO_EXPORT_TRUECOLOR                36
#define WIN_INFO_EXPORT_DITHER                   37
#define WIN_INFO_EXPORT_TRANSPRASTER             38
#define WIN_INFO_EXPORT_TRANSPVECTOR             39
#define WIN_INFO_PRINTER_SCALE_PATTERNS          40
#define WIN_INFO_EXPORT_ANTIALIASING             41
#define WIN_INFO_EXPORT_THRESHOLD                42
#define WIN_INFO_EXPORT_MASKSIZE                 43
#define WIN_INFO_EXPORT_FILTER                   44
#define WIN_INFO_ENHANCED_RENDERING              45
#define WIN_INFO_SMOOTH_TEXT                     46
#define WIN_INFO_SMOOTH_IMAGE                    47
#define WIN_INFO_SMOOTH_VECTOR                   48
#define WIN_INFO_PARENT_LAYOUT                   49
#define WIN_INFO_CLIENTWIDTH                     50
#define WIN_INFO_CLIENTHEIGHT                    51
#define WIN_INFO_PRINTER_RESOLUTION              52
#define WIN_INFO_TOPMOSTWHENFLOATING             53

//----------------------------------------------------------------------------
// Window types, returned by WindowInfo() for WIN_INFO_TYPE
//----------------------------------------------------------------------------
#define WIN_MAPPER                               1
#define WIN_BROWSER                              2
#define WIN_LAYOUT                               3
#define WIN_GRAPH                                4
#define WIN_BUTTONPAD                            19
#define WIN_TOOLBAR                              25
#define WIN_CART_LEGEND                          27
#define WIN_3DMAP                                28
#define WIN_ADORNMENT                            32
#define WIN_LEGEND_DESIGNER                      35
#define WIN_LAYOUT_DESIGNER                      36
#define WIN_CUSTOM                               37
#define WIN_HELP                                 1001
#define WIN_MAPBASIC                             1002
#define WIN_MESSAGE                              1003
#define WIN_RULER                                1007
#define WIN_INFO                                 1008
#define WIN_LEGEND                               1009
#define WIN_STATISTICS                           1010
#define WIN_MAPINFO                              1011

//----------------------------------------------------------------------------
// Version 2 window types no longer used in version 3 or later versions
//----------------------------------------------------------------------------
#define WIN_TOOLPICKER                           1004
#define WIN_PENPICKER                            1005
#define WIN_SYMBOLPICKER                         1006

//----------------------------------------------------------------------------
// Window types which can be used in Open/Close Window statements
//----------------------------------------------------------------------------
#define WIN_TABLE_LIST                           2001
#define WIN_LAYER_CONTROL                        2002
#define WIN_MOVE_MAP_TO                          2003

//----------------------------------------------------------------------------
// Window states, returned by WindowInfo() for WIN_INFO_STATE
//----------------------------------------------------------------------------
#define WIN_STATE_NORMAL                         0
#define WIN_STATE_MINIMIZED                      1
#define WIN_STATE_MAXIMIZED                      2
#define WIN_STATE_DOCKED                         3
#define WIN_STATE_TABBED                         4
#define WIN_STATE_AUTOHIDDEN                     5
#define WIN_STATE_HIDDEN                         6

//----------------------------------------------------------------------------
// Print orientation, returned by WindowInfo() for WIN_INFO_PRINTER_ORIENT
//----------------------------------------------------------------------------
#define WIN_PRINTER_PORTRAIT                     1
#define WIN_PRINTER_LANDSCAPE                    2

//----------------------------------------------------------------------------
// Antialiasing filters, returned by WindowInfo() for WIN_INFO_EXPORT_FILTER
//----------------------------------------------------------------------------
#define FILTER_VERTICALLY_AND_HORIZONTALLY       0
#define FILTER_ALL_DIRECTIONS_1                  1
#define FILTER_ALL_DIRECTIONS_2                  2
#define FILTER_DIAGONALLY                        3
#define FILTER_HORIZONTALLY                      4
#define FILTER_VERTICALLY                        5

//----------------------------------------------------------------------------
// GetPreference() defines
//----------------------------------------------------------------------------

//----------------------------------------------------------------------------
// Preference Segment defines
//----------------------------------------------------------------------------
#define   PREFERENCE_SYSTEM                          0
#define   PREFERENCE_TABLES                          1
#define   PREFERENCE_BROWSER                         3
#define   PREFERENCE_MAPPER                          4
#define   PREFERENCE_INTERNATIONAL                   6
#define   PREFERENCE_PATHS                           11
#define   PREFERENCE_LEGEND                          12
#define   PREFERENCE_OUTPUT                          13
#define   PREFERENCE_PRINTER                         14
#define   PREFERENCE_WEBSERVICES                     15
#define   PREFERENCE_IMAGE_PROCESSING                16
#define   PREFERENCE_LAYOUT                          17
#define   PREFERENCE_PDF                             18
#define   PREFERENCE_PERFORMANCE                     19
#define   PREFERENCE_MANAGED_WORKSPACE               50

//----------------------------------------------------------------------------
// System Preference defines
//----------------------------------------------------------------------------
#define PREF_SYSTEM_COLORDEFAULT                     0
#define PREF_SYSTEM_AUTOSAVEWORKSPACE                3
#define PREF_SYSTEM_COPYTEXTCLIPBOARD                7
#define PREF_SYSTEM_COPYGRAPHICSCLIPBOARD            8
#define PREF_SYSTEM_PAPERUNITS                       9
#define PREF_SYSTEM_UNDO_COUNT                       10
#define PREF_SYSTEM_SUBSTITUTETTFONT                 12
#define PREF_SYSTEM_COPY_EMF_CLIPBOARD               13
#define PREF_SYSTEM_COPY_OLEOBJECT                   14
#define PREF_SYSTEM_SAVEQUERYWORKSPACE               18
#define PREF_SYSTEM_DATEWINDOW                       19
#define PREF_SYSTEM_UNDO_SIZE                        20
#define PREF_SYSTEM_RESOLUTION                       21
#define PREF_SYSTEM_VMOPTION                         22
#define PREF_SYSTEM_DISTANCEUNITS                    23
#define PREF_SYSTEM_AREAUNITS                        24
#define PREF_SYSTEM_COPY_OLEOBJECT_NEW               25
#define PREF_SYSTEM_COPY_ENHANCED_EMF_TYPE           26
#define PREF_SYSTEM_CROSSHAIRCURSOR                  28
#define PREF_SYSTEM_COPY_WMF_CLIPBOARD               29
#define PREF_SYSTEM_TRUECOLORCURSORS                 30
#define PREF_SYSTEM_WRITETOOLWINDOWSTOWORKSPACE      31
#define PREF_SYSTEM_MULTIPLETOUCH                    32
#define PREF_SYSTEM_DEFAULTFILEFORMAT                33
#define PREF_SYSTEM_DEFAULTCHARSETFORNATIVEX         34
#define PREF_SYSTEM_WRITEUNICODEDATA                 35

//----------------------------------------------------------------------------
// Workspace Preference defines
//----------------------------------------------------------------------------
#define PREF_WKS_PROMPTTOSAVE                        1
#define PREF_WKS_PROMPTTOSAVEOPTION                  2

//----------------------------------------------------------------------------
// Table Preference defines
//----------------------------------------------------------------------------
#define PREF_TABLES_FILEOPENWHAT                     0
#define PREF_TABLES_OPENCHOICESCREEN                 1
#define PREF_TABLES_CONNNAME                         2
#define PREF_TABLES_CONNDRIVER                       3
#define PREF_TABLES_CONNSTRING                       4

//----------------------------------------------------------------------------
// Browser Preference defines
//----------------------------------------------------------------------------
#define PREF_BROWSER_RDTSORTORDER                    0
#define PREF_BROWSER_ALTROWCOLORRED                  1
#define PREF_BROWSER_ALTROWCOLORGREEN                2
#define PREF_BROWSER_ALTROWCOLORBLUE                 3

//----------------------------------------------------------------------------
// Mapper Preference defines
//----------------------------------------------------------------------------
#define PREF_MAPPER_PRESERVESCALEONRESIZE            0
#define PREF_MAPPER_HIGHLIGHT_BRUSH                  2
#define PREF_MAPPER_SNAPTONODES                      3
#define PREF_MAPPER_MOVEDUPLICATENODES               4
#define PREF_MAPPER_PIXELTOLERANCE                   5
#define PREF_MAPPER_TARGET_PEN                       7
#define PREF_MAPPER_TARGET_BRUSH                     8
#define PREF_MAPPER_HIGHLIGHT_PEN                    9
#define PREF_MAPPER_SCROLLBARS                       10
#define PREF_MAPPER_SAVECOSMETIC                     11
#define PREF_MAPPER_SAVELABELS                       12
#define PREF_MAPPER_SAVETHEMATIC                     13
#define PREF_MAPPER_SHOWINFOTIPS                     14
#define PREF_MAPPER_DISTANCETYPE                     15
#define PREF_MAPPER_SHOWDEGREESMINUTESSECONDS        16
#define PREF_MAPPER_LASTTHEMETYPE                    17
#define PREF_MAPPER_GRIDOUTPUTFORMATHANDLER          18
#define PREF_MAPPER_CLIPREGIONTYPE                   19
#define PREF_MAPPER_3DHARDWARE                       20
#define PREF_MAPPER_AUTORASTERZOOMLAYER              21
#define PREF_MAPPER_AUTOGRIDZOOMLAYER                22
#define PREF_MAPPER_POINT_HIGHLIGHT_BRUSH            23
#define PREF_MAPPER_POINT_TARGET_BRUSH               24
#define PREF_MAPPER_AUTOSCROLL                       25
#define PREF_MAPPER_DEFAULTPROJECTION                26
#define PREF_MAPPER_DEFAULTMBPROJECTION              27
#define PREF_MAPPER_DISPLAYSNAPRADIUS                28
#define PREF_MAPPER_CARTOGRAPHICSCALE                29
#define PREF_MAPPER_DISPLAYLAYERCONTROLTOOLTIPS      30
#define PREF_MAPPER_AUTONODETOLERANCE                31
#define PREF_MAPPER_DRAWLAYERSUNDERTHEMES            32
#define PREF_MAPPER_READDEFAULTTHEMEMETADATA         33
#define PREF_MAPPER_FINDSELECTIONAFTERPASTE          34
#define PREF_MAPPER_ZOOMONFINDSELECTION              35
#define PREF_MAPPER_ALPHAHIGHLIGHT                   36
#define PREF_MAPPER_ZOOMMOUSEWHEELONCURSOR           37
#define PREF_MAPPER_ADORNAUTOMATICSCALING            38
#define PREF_MAPPER_ADORN_BACKGROUND_BRUSH           39
#define PREF_MAPPER_ADORN_BACKGROUND_PEN             40

//----------------------------------------------------------------------------
// International Preference defines
//----------------------------------------------------------------------------
#define PREF_INTERNATIONAL_ADDRMATCH                 0
#define PREF_INTERNATIONAL_INSET                     1
#define PREF_INTERNATIONAL_OFFSET                    2
#define PREF_INTERNATIONAL_DISTUNIT                  3
#define PREF_INTERNATIONAL_ENABLEDISTANCESPECIFIED   4
#define PREF_INTERNATIONAL_INSET_DIST                5
#define PREF_INTERNATIONAL_INSET_DIST_DISTUNIT       6
#define PREF_INTERNATIONAL_USER_CLEAN                7
#define PREF_INTERNATIONAL_CLOSE_ADDR                8
#define PREF_INTERNATIONAL_ALT_BDY                   9

//----------------------------------------------------------------------------
// Path Preference defines
//----------------------------------------------------------------------------
#define PREF_PATH_MITABLEPATH                        0
#define PREF_PATH_MIWORKSPACEPATH                    1
#define PREF_PATH_MIPROGRAMPATH                      2
#define PREF_PATH_MIDATAPATHS                        3
#define PREF_PATH_USETRAVERSAL                       4
#define PREF_PATH_CUSTOMSYMBOLPATH                   5
#define PREF_PATH_IMPORTPATH                         6
#define PREF_PATH_SQLQUERYPATH                       7
#define PREF_PATH_THEMETEMPLATEPATH                  8
#define PREF_PATH_MIQUERYPATH                        9
#define PREF_PATH_NEWGRIDPATH                        10
#define PREF_PATH_REMOTETABLEPATH                    13
#define PREF_PATH_TABLESFROMSHAPEFILES               14
#define PREF_PATH_TABLESFROMWFS                      15
#define PREF_PATH_TABLESFROMWMS                      16
#define PREF_PATH_LAYOUTIMAGES                       17

//----------------------------------------------------------------------------
// Legend Preference defines
//----------------------------------------------------------------------------
#define PREF_LEGEND_FNTTITLE                         0
#define PREF_LEGEND_FNTSUBTITLE                      1
#define PREF_LEGEND_FNTSTYLE                         2
#define PREF_LEGEND_PENBORDER                        3
#define PREF_LEGEND_TITLEPATTERN                     4
#define PREF_LEGEND_SUBTITLEPATTERN                  5
#define PREF_LEGEND_STYLEPATTERN                     6
#define PREF_LEGEND_STYLESIZE                        7
#define PREF_LEGEND_PROMPTSAVEWORKSPACE              8
#define PREF_LEGEND_CREATEPRE1150LEGENDS             9
#define PREF_LEGEND_PAPERUNITS                       10
#define PREF_LEGEND_SYNCDISPLAY                      11
#define PREF_LEGEND_LINEWIDTH                        12
#define PREF_LEGEND_REGIONWIDTH                      13
#define PREF_LEGEND_REGIONHEIGHT                     14
#define PREF_LEGEND_DISPLAYGRID                      15
#define PREF_LEGEND_SNAPGRID                         16
#define PREF_LEGEND_AUTOUPDATE                       17
#define PREF_LEGEND_GRIDCOLORALPHA                   18
#define PREF_LEGEND_GRIDCOLORRED                     19
#define PREF_LEGEND_GRIDCOLORGREEN                   20
#define PREF_LEGEND_GRIDCOLORBLUE                    21
#define PREF_LEGEND_GUIDECOLORALPHA                  22
#define PREF_LEGEND_GUIDECOLORRED                    23
#define PREF_LEGEND_GUIDECOLORGREEN                  24
#define PREF_LEGEND_GUIDECOLORBLUE                   25
#define PREF_LEGEND_AUTOFONTSIZESETTING              26

//----------------------------------------------------------------------------
// Output Preference defines
//----------------------------------------------------------------------------
#define PREF_OUTPUT_USETRUECOLORDISPLAY              0
#define PREF_OUTPUT_DITHERMETHODDISPLAY              1
#define PREF_OUTPUT_PRINTMETHOD                      2
#define PREF_OUTPUT_PRINTBORDER                      3
#define PREF_OUTPUT_TRANSPWHENPRINTVECTOR            5
#define PREF_OUTPUT_USETRUECOLORPRINT                6
#define PREF_OUTPUT_DITHERMETHODPRINT                7
#define PREF_OUTPUT_EXPORTBORDER                     8
#define PREF_OUTPUT_TRANSPWHENEXPORTRASTER           9
#define PREF_OUTPUT_TRANSPWHENEXPORTVECTOR           10
#define PREF_OUTPUT_USETRUECOLOREXPORT               11
#define PREF_OUTPUT_DITHERMETHODEXPORT               12
#define PREF_OUTPUT_ROPTRANSPWHENPRINTINGRASTER      13
#define PREF_OUTPUT_ROPTRANSPWHENEXPORTINGRASTER     14
#define PREF_OUTPUT_SCALEPATTERNSPRINT               15
#define PREF_OUTPUT_USEANTIALIASING                  16
#define PREF_OUTPUT_ANTIALIASINGTHRESHOLD            17
#define PREF_OUTPUT_ANTIALIASINGMASKSIZE             18
#define PREF_OUTPUT_ANTIALIASINGFILTER               19
#define PREF_OUTPUT_DISPLAYPOLYGONMODE               20
#define PREF_OUTPUT_PRINTINGPOLYGONMODE              21
#define PREF_OUTPUT_EXPORTINGPOLYGONMODE             22
#define PREF_OUTPUT_SUBDIVIDEPRINTING                23
#define PREF_OUTPUT_CLARIFYTEXTOVERWYSIWYG           24
#define PREF_OUTPUT_DISPLAYPOINTSFAST                25

//----------------------------------------------------------------------------
// Printer Preference defines
//----------------------------------------------------------------------------
#define PREF_PRINTER_USESPECIFICPRINTER              0
#define PREF_PRINTER_PAPERSIZE                       1
#define PREF_PRINTER_ORIENTATION                     2
#define PREF_PRINTER_PRINTERNAME                     3
#define PREF_PRINTER_SAVEPRINTERINTOWORKSPACE        4
#define PREF_PRINTER_RESTOREPRINTERFROMWORKSPACE     5

//----------------------------------------------------------------------------
// WebServices Preference defines
//----------------------------------------------------------------------------
#define PREF_WEBSERVICES_WFSCONNECTTIMEOUT           0
#define PREF_WEBSERVICES_WFSSENDTIMEOUT              1
#define PREF_WEBSERVICES_WFSRECEIVETIMEOUT           2
#define PREF_WEBSERVICES_WMSCONNECTTIMEOUT           3
#define PREF_WEBSERVICES_WMSSENDTIMEOUT              4
#define PREF_WEBSERVICES_WMSRECEIVETIMEOUT           5
#define PREF_WEBSERVICES_WMSMAXWIDTH                 6
#define PREF_WEBSERVICES_WMSMAXHEIGHT                7
#define PREF_WEBSERVICES_ACCESSTYPE                  8
#define PREF_WEBSERVICES_PROXYSERVER                 9
#define PREF_WEBSERVICES_PORT                        10
#define PREF_WEBSERVICES_BYPASSLOCAL                 11
#define PREF_WEBSERVICES_BYPASSLIST                  12
#define PREF_WEBSERVICES_WFSREFRESHUSINGMAP          13
#define PREF_WEBSERVICES_WEBGEOCODECONNECTTIMEOUT    14
#define PREF_WEBSERVICES_WEBGEOCODESENDTIMEOUT       15
#define PREF_WEBSERVICES_WEBGEOCODERECEIVETIMEOUT    16
#define PREF_WEBSERVICES_STREETOFFSET                17
#define PREF_WEBSERVICES_CORNEROFFSET                18
#define PREF_WEBSERVICES_UNITOFFSET                  19
#define PREF_WEBSERVICES_GEOCODEDICTIONARY           20
#define PREF_WEBSERVICES_GEOCODEBATCHSIZE            21
#define PREF_WEBSERVICES_WEBROUTINGCONNECTTIMEOUT    22
#define PREF_WEBSERVICES_WEBROUTINGSENDTIMEOUT       23
#define PREF_WEBSERVICES_WEBROUTINGRECEIVETIMEOUT    24
#define PREF_WEBSERVICES_DRIVEREGIONBATCHSIZE        25
#define PREF_WEBSERVICES_APPLYMULTIRESULTCODE        26
#define PREF_WEBSERVICES_READTIMEOUTWEBTILESERVER    27
#define PREF_WEBSERVICES_REQUESTTIMEOUTWEBTILESERVER 28
#define PREF_WEBSERVICES_LIBRARYSERVICEACCESSMODE    29
#define PREF_WEBSERVICES_OPENMANAGEDTABLESEDITOR     30
#define PREF_WEBSERVICES_BINGKEYFROMUSER             31

//----------------------------------------------------------------------------
// Image Processing Preference defines
//----------------------------------------------------------------------------
#define PREF_IMAGEPROCESSING_REPROJECTIONUSEINDEX    0
#define PREF_IMAGEPROCESSING_RESAMPLEMETHODINDEX     1

//----------------------------------------------------------------------------
// Layout Preference defines
//----------------------------------------------------------------------------
#define PREF_LAYOUT_PROMPTSAVEWORKSPACE              3

//----------------------------------------------------------------------------
// Pdf Preference defines
//----------------------------------------------------------------------------
#define PREF_PDF_LAYEREDPDF                          0
#define PREF_PDF_SAVEGEOINFO                         1
#define PREF_PDF_SAVEATTRIBUTEDATA                   2
#define PREF_PDF_AUTOOPENPDF                         3
#define PREF_PDF_POSTSCRIPTSIM                       4
#define PREF_PDF_SAVEGEOINFOINNATIVEUNITS            5

//----------------------------------------------------------------------------
// Performance Preference defines
//----------------------------------------------------------------------------
#define PREF_PERFORMANCE_CONCURRENCYSETTING          0
#define PREF_PERFORMANCE_DROPINDEXSETTING            1

//============================================================================
// end of MAPBASIC.DEF
//============================================================================

//============================================================================
// MapInfo version 17.0.3 - Menu Item Definitions
//----------------------------------------------------------------------------
// This file contains defines useful when programming in the MapBasic
// language.  The definitions in this file describe the standard MapInfo
// functionality available via the "Run Menu Command" MapBasic statement.
//----------------------------------------------------------------------------
// The defines in this file are organized to match the sequence of
// declarations in the MAPINFOW.MNU file, which in turn reflects the
// organization of the MapInfo menus and buttonpads.
//============================================================================

//----------------------------------------------------------------------------
// File menu
//----------------------------------------------------------------------------
#define M_FILE_NEW                              101
#define M_FILE_OPEN                             102
#define M_FILE_OPEN_WMS                         118
#define M_FILE_OPEN_WFS                         119
#define M_FILE_OPEN_WMTS                        120
#define M_FILE_OPEN_ODBC_CONN                   125
#define M_FILE_OPEN_UNIVERSAL_DATA              126
#define M_FILE_OPEN_BING_AERIAL                 128
#define M_FILE_OPEN_BING_HYBRID                 129
#define M_FILE_OPEN_BING_ROADS                  130
//NOTE: M_FILE_OPEN_OSM_ROADS  was deprecated in 12.0
//#define M_FILE_OPEN_OSM_ROADS                 131
#define M_FILE_OPEN_LASTDLG                     132
#define M_FILE_OPEN_GEOMAP_BRONZE               133
#define M_FILE_OPEN_GEOMAP_STEEL                134
#define M_FILE_OPEN_GEOMAP_IRON                 135
#define M_FILE_CLOSE                            103
#define M_FILE_CLOSE_ALL                        104
#define M_FILE_CLOSE_ODBC                       124
#define M_FILE_SAVE                             105
#define M_FILE_SAVE_COPY_AS                     106
#define M_FILE_SAVE_QUERY                       117
#define M_FILE_SAVE_WORKSPACE                   109
#define M_FILE_SAVE_CURRENT_WORKSPACE           136
#define M_FILE_SAVE_WINDOW_AS                   609
#define M_FILE_REVERT                           107
#define M_FILE_PAGE_SETUP                       111
#define M_FILE_PRINT                            112
#define M_FILE_PRINT_TO_PDF                     127
#define M_FILE_EXIT                             113
#define M_SERVER_LIST                           236

//----------------------------------------------------------------------------
// Edit menu
//----------------------------------------------------------------------------
#define M_EDIT_UNDO                             201
#define M_EDIT_CUT                              202
#define M_EDIT_COPY                             203
#define M_EDIT_PASTE                            204
#define M_EDIT_CLEAR                            205
#define M_EDIT_CLEAROBJ                         206
#define M_EDIT_RESHAPE                          1601
#define M_EDIT_NEW_ROW                          702
#define M_EDIT_GETINFO                          207

//----------------------------------------------------------------------------
// Tools menu
//----------------------------------------------------------------------------
#define M_TOOLS_CRYSTAL_REPORTS_NEW             1802
#define M_TOOLS_CRYSTAL_REPORTS_OPEN            1803
#define M_TOOLS_RUN                             110
#define M_MBTOOL_GET_MB_UTILITIES               1811
#define M_TOOLS_TOOL_MANAGER                    1801

//----------------------------------------------------------------------------
// Objects menu
//----------------------------------------------------------------------------
#define M_OBJECTS_SET_TARGET                    1610
#define M_OBJECTS_CLEAR_TARGET                  1611
#define M_OBJECTS_COMBINE                       1605
#define M_OBJECTS_DISAGG                        1621
#define M_OBJECTS_BUFFER                        1606
#define M_OBJECTS_CONVEX_HULL                   1616
#define M_OBJECTS_ENCLOSE                       1617
#define M_OBJECTS_VORONOI                       1622
#define M_OBJECTS_DRIVE_REGION                  1627
#define M_OBJECTS_SPLIT                         1612
#define M_OBJECTS_ERASE                         1613
#define M_OBJECTS_ERASE_OUT                     1614
#define M_OBJECTS_POLYLINE_SPLIT                1623
#define M_OBJECTS_POLYLINE_SPLIT_AT_NODE        1626
#define M_OBJECTS_OVERLAY                       1615
#define M_OBJECTS_CHECK_REGIONS                 1618
#define M_OBJECTS_CLEAN                         1619
#define M_OBJECTS_SNAP                          1620
#define M_OBJECTS_OFFSET                        1624
#define M_OBJECTS_ROTATE                        1625
#define M_OBJECTS_SMOOTH                        1602
#define M_OBJECTS_UNSMOOTH                      1603
#define M_OBJECTS_CVT_PGON                      1607
#define M_OBJECTS_CVT_PLINE                     1604

//----------------------------------------------------------------------------
// Query menu
//----------------------------------------------------------------------------
#define M_QUERY_SELECT                          301
#define M_QUERY_SQLQUERY                        302
#define M_QUERY_SELECTALL                       303
#define M_QUERY_INVERTSELECT                    311
#define M_QUERY_UNSELECT                        304
#define M_QUERY_FIND                            305
#define M_QUERY_FIND_SELECTION_CURRENT_MAP      312
#define M_QUERY_FIND_SELECTION                  306
#define M_QUERY_FIND_ADDRESS                    313
#define M_QUERY_CALC_STATISTICS                 309

//----------------------------------------------------------------------------
// Table, Maintenance, and Raster menus
//----------------------------------------------------------------------------
#define M_TABLE_UPDATE_COLUMN                   405
#define M_TABLE_APPEND                          411
#define M_TABLE_GEOCODE                         407
#define M_TABLE_WEB_GEOCODE                     424
#define M_TABLE_CREATE_POINTS                   408
#define M_TABLE_MERGE_USING_COLUMN              406
#define M_TABLE_BUFFER                          419
#define M_TABLE_VORONOI                         420
#define M_TABLE_DRIVE_REGION                    425
#define M_TABLE_IMPORT                          401
#define M_TABLE_EXPORT                          402

#define M_TABLE_MODIFY_STRUCTURE                404
#define M_TABLE_DELETE                          409
#define M_TABLE_RENAME                          410
#define M_TABLE_PACK                            403
#define M_TABLE_MAKEMAPPABLE                    415
#define M_TABLE_CHANGESYMBOL                    418
#define M_TABLE_UNLINK                          416
#define M_TABLE_REFRESH                         417

#define M_ORACLE_CREATE_WORKSPACE               1804
#define M_ORACLE_DELETE_WORKSPACE               1805
#define M_ORACLE_VERSION_ENABLE_ON              1806
#define M_ORACLE_VERSION_ENABLE_OFF             1807
#define M_ORACLE_MERGE_PARENT                   1808
#define M_ORACLE_REFRESH_FROM_PARENT            1809

#define M_TABLE_RASTER_STYLE                    414
#define M_TABLE_RASTER_REG                      413
#define M_TOOLS_RASTER_REG                      1730

#define M_TABLE_WMS_PROPS                       421
#define M_TABLE_WFS_REFRESH                     422
#define M_TABLE_WFS_PROPS                       423
#define M_TABLE_UNIVERSAL_DATA_REFRESH          426

//----------------------------------------------------------------------------
// Options and Preferences menus
//----------------------------------------------------------------------------
#define M_FORMAT_PICK_LINE                      501
#define M_FORMAT_PICK_FILL                      502
#define M_FORMAT_PICK_SYMBOL                    503
#define M_FORMAT_PICK_FONT                      504
#define M_WINDOW_BUTTONPAD                      605
#define M_WINDOW_LEGEND                         606
#define M_WINDOW_STATISTICS                     607
#define M_WINDOW_MAPBASIC                       608
#define M_WINDOW_STATUSBAR                      616
#define M_FORMAT_CUSTOM_COLORS                  617
#define M_EDIT_PREFERENCES                      208

#define M_EDIT_PREFERENCES_SYSTEM               210
#define M_EDIT_PREFERENCES_MAP                  212
#define M_EDIT_PREFERENCES_LEGEND               215
#define M_EDIT_PREFERENCES_FILE                 211
#define M_EDIT_PREFERENCES_COUNTRY              213
#define M_EDIT_PREFERENCES_PATH                 214
#define M_EDIT_PREFERENCES_OUTPUT               216
#define M_EDIT_PREFERENCES_PRINTER              217
#define M_EDIT_PREFERENCES_STYLES               218
#define M_EDIT_PREFERENCES_IMAGE_PROC           219
#define M_EDIT_PREFERENCES_WEBSERVICES          220
#define M_EDIT_PREFERENCES_LAYOUT               221
#define M_EDIT_PREFERENCES_BROWSER              222
#define M_EDIT_PREFERENCES_NOTIFICATIONS        223
#define M_EDIT_PREFERENCES_PERFORMANCE          224
#define M_EDIT_PREFERENCES_WORKSPACE            225
#define M_EDIT_PREFERENCES_MAPBASIC             226
#define M_EDIT_PREFERENCES_EXPLORER             227
#define M_EDIT_PREFERENCES_WINDOWSTATE          228
#define M_EDIT_PREFERENCES_TASKMANAGER          229
#define M_EDIT_PREFERENCES_WINDOWLIST           230
#define M_EDIT_PREFERENCE_RESETEXPLORERHEIGHTS  231
#define M_EDIT_PREFERENCE_SAVERIBBONSTATE       232
#define M_EDIT_PREFERENCE_DELETERIBBONSTATE     233
#define M_EDIT_PREFERENCE_SAVETOOLWINDOWSTATE   234
#define M_EDIT_PREFERENCE_DELETETOOLWINDOWSTATE 235

//----------------------------------------------------------------------------
// Window menu
//----------------------------------------------------------------------------
#define M_WINDOW_BROWSE                         601
#define M_WINDOW_MAP                            602
#define M_WINDOW_GRAPH                          603
#define M_WINDOW_LAYOUT                         604
#define M_WINDOW_LAYOUT_DESIGNER                620
#define M_WINDOW_REDISTRICT                     615
#define M_WINDOW_REDRAW                         610
#define M_WINDOW_TILE                           611
#define M_WINDOW_CASCADE                        612
#define M_WINDOW_ARRANGEICONS                   613
#define M_WINDOW_MORE                           614
#define M_WINDOW_FIRST                          625
#define M_WINDOW_RECOVER_FLOAT                  619
#define M_WINDOW_MESSAGE_CLEAR                  621
#define M_WINDOW_CLONE                          622

// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
// Note: the 2nd through 75th windows can be accessed as (M_WINDOW_FIRST+i-1)
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

//----------------------------------------------------------------------------
// Help menu
//----------------------------------------------------------------------------
#define M_HELP_CONTEXTSENSITIVE                 1201
#define M_HELP_CONTENTS                         1202

#define M_HELP_SEARCH                           1203
#define M_HELP_MAPINFO_3DGRAPH_HELP             1214
#define M_HELP_MAPINFO_WWW                      1210
// NOTE: M_HELP_MAPINFO_CONNECT_SERVICES was deprecated in 12.0
// #define M_HELP_MAPINFO_CONNECT_SERVICES         1215
#define M_HELP_MAPINFO_WWW_TUTORIAL             1213
#define M_HELP_MAPINFO_WWW_STORE                1211
#define M_HELP_CONNECT_MIFORUM                  1209
#define M_HELP_CHECK_FOR_UPDATE                 1212
#define M_HELP_MAPINFO_PRODUCT_FEEDBACK         1219
#define M_HELP_BING_TERMS_OF_USE                1220
#define M_HELP_MAPINFO_NEWSFEED                 1221

#define M_HELP_MAPINFO_RETURN_LICENSE           1216
#define M_HELP_MAPINFO_BORROW_LICENSE           1217
#define M_HELP_MAPINFO_RETURN_BORROWED_LICENSE  1218
#define M_HELP_MAPINFO_USER_BING_KEY            1222
#define M_HELP_MAPINFO_START_ADVANCED_EVALUATION     1223
#define M_HELP_MAPINFO_FINISH_ADVANCED_EVALUATION    1224

#define M_HELP_ABOUT                            1205

#define M_HELP_USE_HELP                         1204
#define M_HELP_HELPMODE                         1206
#define M_HELP_TECHSUPPORT                      1208

//----------------------------------------------------------------------------
// Browse menu
//----------------------------------------------------------------------------
#define M_BROWSE_PICK_FIELDS                    704
#define M_BROWSE_OPTIONS                        703
#define M_BROWSE_SORT                           710
#define M_BROWSE_TOOLBAR                        711
#define M_BROWSE_HOTLINK_OPTIONS                714

// Browse window context menu
#define M_BROWSE_SORT_ASCENDING                 2702
#define M_BROWSE_SORT_DESCENDING                2703
#define M_BROWSE_FILTER                         2704
#define M_BROWSE_CLEAR_FILTER                   2705

// Browse window Sort/Filter menu (via toolbar)
#define M_BROWSE_REAPPLY_SORT_FILTERS           2709
#define M_BROWSE_TOGGLE_SORT_FILTERS            2706
#define M_BROWSE_CLEAR_SORT_FILTERS             2710
#define M_BROWSE_CLEAR_FILTERS                  2707
#define M_BROWSE_CLEAR_SORT                     2701
#define M_BROWSE_CREATE_QUERY                   2711

// Browse window Selection menu (via toolbar)
#define M_BROWSE_SELECT_ALL_IN_BROWSER          2708

//----------------------------------------------------------------------------
// Map menu
//----------------------------------------------------------------------------
// Open Layer Control in a window:
#define M_MAP_LAYER_CONTROL                     822
// Open Layer Control in a dialog box:
#define M_MAP_LAYER_CONTROL_DIALOG              801
#define M_MAP_CREATE_3DMAP                      817
#define M_MAP_CREATE_PRISMMAP                   818
#define M_MAP_THEMATIC                          307
#define M_MAP_MODIFY_THEMATIC                   308
#define M_MAP_CREATE_LEGEND                     816
#define M_MAP_CREATE_SCALEBAR                   824
#define M_MAP_CHANGE_VIEW                       805
#define M_MAP_CLONE_MAPPER                      811
#define M_MAP_PREVIOUS                          806
#define M_MAP_ENTIRE_LAYER                      807
#define M_MAP_CLEAR_CUSTOM_LABELS               814
#define M_MAP_SAVE_COSMETIC                     809
#define M_MAP_CLEAR_COSMETIC                    810
#define M_MAP_SET_CLIP_REGION                   812
#define M_MAP_CLIP_REGION_ONOFF                 813
#define M_MAP_SETUPDIGITIZER                    803
#define M_MAP_OPTIONS                           802
#define M_MAP_MOVE_MAP_TO                       825
#define M_MAP_CREATE_LEGEND_PREF                826
#define M_MAP_CREATE_LEGEND_DESIGNER            827
#define M_VIEW_ZOOM_TO_TILESERVER_LAYER         828
#define M_VIEW_WORKSPACE_EXPLORER               829
#define M_VIEW_WINDOW_LIST                      830
#define M_VIEW_TOOL_MANAGER                     831
#define M_VIEW_TASK_MANAGER                     832
#define M_VIEW_CONNECTION_LIST                  833
#define M_VIEW_INFO_WINDOW                      834
#define M_VIEW_MESSAGE_WINDOW                   835
#define M_VIEW_STATISTICS_WINDOW                836
#define M_VIEW_RULER_WINDOW                     837
#define M_VIEW_LOCK_SCALE                       838
#define M_VIEW_SELECTABLE_LABELS                839
#define M_VIEW_ENTIRE_MAP                       840
#define M_VIEW_ZOOM_TO_SELECTED_OBJECTS         841
#define M_VIEW_SNAP_TO_NODE                     842
#define M_VIEW_SUSPEND_REDRAW                   843
#define M_VIEW_LEGEND_WINDOW                    844

// Map window context menu
#define M_MAP_AUTOSCROLL_ONOFF                  815
#define M_MAP_VIEW_ENTIRE_SELECTION             819
#define M_MAP_SELECT_ALL_FROM_SELECTION         820
#define M_MAP_MAKE_SELECTION_EDITABLE           821

//----------------------------------------------------------------------------
// Layout menu
//----------------------------------------------------------------------------
#define M_LAYOUT_CHANGE_VIEW                    902
#define M_LAYOUT_ACTUAL                         903
#define M_LAYOUT_ENTIRE                         904
#define M_LAYOUT_PREVIOUS                       905
#define M_LAYOUT_BRING2FRONT                    906
#define M_LAYOUT_SEND2BACK                      907
#define M_LAYOUT_ALIGN                          908
#define M_LAYOUT_DROPSHADOWS                    909
#define M_LAYOUT_DISPLAYOPTIONS                 901
#define M_LAYOUT_AUTOSCROLL_ONOFF               910
#define M_LAYOUT_ALIGN_LEFT                     911
#define M_LAYOUT_ALIGN_RIGHT                    912
#define M_LAYOUT_ALIGN_TOP                      913
#define M_LAYOUT_ALIGN_BOTTOM                   914

//----------------------------------------------------------------------------
// Layout Designer menu
//----------------------------------------------------------------------------
#define M_LAYOUT_NEW_MAP                       2901
#define M_LAYOUT_NEW_BROWSER                   2902
#define M_LAYOUT_PAGE_SETUP                    2903
#define M_LAYOUT_FIT_IN_WINDOW                 2904
#define M_LAYOUT_ITEM_ACTIVATE                 2905
#define M_LAYOUT_BRING_TO_FRONT                2906
#define M_LAYOUT_SEND_TO_BACK                  2907
#define M_LAYOUT_FRAME_PROP                    2908
#define M_LAYOUT_TEXT_STYLE                    2909
#define M_LAYOUT_ADD_LEGENDS                   2910
#define M_LAYOUT_REFRESH_LEGENDS               2911
#define M_LAYOUT_MODIFY_THEMATIC               2912
#define M_LAYOUT_REMOVE_FRAME_CONTENT          2913
#define M_LAYOUT_NEW_EMPTY_FRAME               2914
#define M_LAYOUT_REMOVE_SELECTED_ITEMS         2915
#define M_UNDO                                 2916
#define M_REDO                                 2917
#define M_LAYOUT_ADD_HORIZ_GUIDE               2919
#define M_LAYOUT_ADD_SYMBOL                    2921
#define M_LAYOUT_ADD_LINE                      2922
#define M_LAYOUT_ADD_ELLIPSE                   2923
#define M_LAYOUT_ADD_RECT                      2924
#define M_LAYOUT_ADD_ROUNDRECT                 2925
#define M_LAYOUT_ADD_VERT_GUIDE                2926
#define M_LAYOUT_CHANGE_FRAME_STYLE            2927
#define M_LAYOUT_TOGGLE_GUIDES                 2928
#define M_LAYOUT_ZOOM_TO_HUNDRED_PERCENT       2929
#define M_LAYOUT_TOGGLE_SMART_GUIDES           2930
#define M_LAYOUT_INSERT_IMAGE                  2931
#define M_LAYOUT_INSERT_TEXT                   2932
#define M_LAYOUT_LINESPACING_DOUBLE            2933
#define M_LAYOUT_LINESPACING_ONEDOTFIVE        2934
#define M_LAYOUT_LINESPACING_SINGLE            2935
#define M_LAYOUT_LINE_STYLE                    2936
#define M_LAYOUT_SYMBOL_STYLE                  2937
#define M_LAYOUT_PAN                           2938
#define M_LAYOUT_REGION_STYLE                  2939
#define M_LAYOUT_REMOVE_GUIDELINES             2940
#define M_LAYOUT_SELECT                        2941
#define M_LAYOUT_SNAP_TO_BOTTOM                2942
#define M_LAYOUT_SNAP_TO_GRID                  2943
#define M_LAYOUT_SNAP_TO_LEFT                  2944
#define M_LAYOUT_SNAP_TO_RIGHT                 2945
#define M_LAYOUT_SNAP_TO_TOP                   2946
#define M_LAYOUT_TEXT_JUSTIFY_CENTER           2947
#define M_LAYOUT_TEXT_JUSTIFY_LEFT             2948
#define M_LAYOUT_TEXT_JUSTIFY_RIGHT            2949
#define M_LAYOUT_TOGGLE_GRIDLINES              2950
#define M_LAYOUT_ZOOM_IN                       2951
#define M_LAYOUT_ZOOM_OUT                      2952
#define M_LAYOUT_ALIGN_HORIZ_CENTER            2953
#define M_LAYOUT_ALIGN_VERT_CENTER             2954
#define M_LAYOUT_ROTATE_FRAME_0                2955
#define M_LAYOUT_ROTATE_FRAME_90               2956
#define M_LAYOUT_ROTATE_FRAME_270              2957
#define M_LAYOUT_ADD_POLYGON                   2958
#define M_LAYOUT_ADD_POLYLINE                  2959
#define M_LAYOUT_DIST_HORIZ_PAGE               2960
#define M_LAYOUT_DIST_VERT_PAGE                2961
#define M_LAYOUT_DIST_HORIZ                    2962
#define M_LAYOUT_DIST_VERT                     2963

//----------------------------------------------------------------------------
// Gridline Menu
//----------------------------------------------------------------------------
#define M_VIEW_GRIDLINE_CREATE                 845
#define M_VIEW_GRIDLINE_PROPERTIES             846

//----------------------------------------------------------------------------
// Graph menu
//----------------------------------------------------------------------------
#define M_GRAPH_TYPE                            1001
#define M_GRAPH_LABEL_AXIS                      1002
#define M_GRAPH_VALUE_AXIS                      1003
#define M_GRAPH_SERIES                          1004

//----------------------------------------------------------------------------
// New Graph menu
//----------------------------------------------------------------------------
#define M_GRAPH_FORMATING                       2007
#define M_GRAPH_GENERAL_OPTIONS                 2002
#define M_GRAPH_SERIES_OPTIONS                  2003
#define M_GRAPH_GRID_SCALES                     2004
#define M_GRAPH_TITLES                          2005
#define M_GRAPH_3D_VIEWING_ANGLE                2006
#define M_GRAPH_SAVE_AS_TEMPLATE                2008

//----------------------------------------------------------------------------
// MapBasic menu
//----------------------------------------------------------------------------
#define M_MAPBASIC_CLEAR                        1101
#define M_MAPBASIC_SAVECONTENTS                 1102

//----------------------------------------------------------------------------
// Redistrict menu
//----------------------------------------------------------------------------
#define M_REDISTRICT_ASSIGN                     705
#define M_REDISTRICT_TARGET                     706
#define M_REDISTRICT_ADD                        707
#define M_REDISTRICT_DELETE                     708
#define M_REDISTRICT_OPTIONS                    709

//----------------------------------------------------------------------------
// Legend menu
//----------------------------------------------------------------------------
#define M_LEGEND_ADD_FRAMES                     1903
#define M_LEGEND_REFRESH                        1902
#define M_LEGEND_PROPERTIES                     1901
#define M_LEGEND_DELETE                         1904

//----------------------------------------------------------------------------
// 3DMap menu
//----------------------------------------------------------------------------
#define M_3DMAP_REFRESH_GRID_TEXTURE            2103
#define M_3DMAP_CLONE_VIEW                      2105
#define M_3DMAP_PREVIOUS_VIEW                   2106
#define M_3DMAP_VIEW_ENTIRE_GRID                2101
#define M_3DMAP_VIEWPOINT_CONTROL               2107
#define M_3DMAP_WIREFRAME                       2104
#define M_3DMAP_PROPERTIES                      2102

//----------------------------------------------------------------------------
// Main Buttonpad
//----------------------------------------------------------------------------
#define M_TOOLS_SELECTOR                        1701
#define M_TOOLS_SEARCH_RECT                     1722
#define M_TOOLS_SEARCH_RADIUS                   1703
#define M_TOOLS_SEARCH_BOUNDARY                 1704
#define M_TOOLS_EXPAND                          1705
#define M_TOOLS_SHRINK                          1706
#define M_TOOLS_RECENTER                        1702
#define M_TOOLS_PNT_QUERY                       1707
#define M_TOOLS_HOTLINK                         1736
#define M_TOOLS_LABELER                         1708
#define M_TOOLS_DRAGWINDOW                      1734
#define M_TOOLS_RULER                           1710

//----------------------------------------------------------------------------
// Drawing Buttonpad
//----------------------------------------------------------------------------
#define M_TOOLS_POINT                           1711
#define M_TOOLS_LINE                            1712
#define M_TOOLS_POLYLINE                        1713
#define M_TOOLS_ARC                             1716
#define M_TOOLS_POLYGON                         1714
#define M_TOOLS_ELLIPSE                         1715
#define M_TOOLS_RECTANGLE                       1717
#define M_TOOLS_ROUNDEDRECT                     1718
#define M_TOOLS_TEXT                            1709
#define M_TOOLS_FRAME                           1719
#define M_TOOLS_ADD_NODE                        1723

//----------------------------------------------------------------------------
// DBMS Buttonpad
//  -- most DBMS buttons defined under M_TABLE_{...} except for this one:
//----------------------------------------------------------------------------
#define M_DBMS_OPEN_ODBC                        116

//----------------------------------------------------------------------------
// Adornment context menu items
//----------------------------------------------------------------------------
#define M_ADORNMENT_PROPERTIES                  2401
#define M_ADORNMENT_REMOVE                      2402

//----------------------------------------------------------------------------
// Layer Control menuitems
//----------------------------------------------------------------------------
#define M_LC_VIEW_ENTIRE                        2501
#define M_LC_LAYER_PROPERTIES                   2502
#define M_LC_LABEL_FONT                         2503
#define M_LC_OTHERS_NO_SELECT                   2504
#define M_LC_RENAME_MAP                         2505
#define M_LC_MAP_OPTIONS                        2506
#define M_LC_ADD_GROUP                          2507
#define M_LC_RENAME_GROUP                       2508
#define M_LC_UNGROUP                            2509
#define M_LC_ADDLAYERS                          2510
#define M_LC_ADD_STYLE_OVERRIDE                 2511
#define M_LC_ADD_LABEL_OVERRIDE                 2512
#define M_LC_DISPLAY_OVERRIDE_PROP              2513
#define M_LC_LABEL_OVERRIDE_PROP                2514
#define M_LC_SPLIT_OVERRIDE                     2515
#define M_LC_REMOVE_OVERRIDE                    2516
#define M_LC_BROWSE_TABLE                       2517
#define M_LC_VIEW_META_DATA                     2518
#define M_LC_ADD_TO_LIBRARY                     2519
#define M_LC_ZOOM_TO_TILESERVER                 2520
#define M_LC_CLEAR_CUSTOM_LABELS                2521
#define M_LC_LABEL_PRIORITIES                   2522
#define M_LC_SELECT_ALL_FROM_LAYER              2523
#define M_LC_OPTIONS                            2524
#define M_LC_LAYER_FRIENDLYNAME                 2542

//----------------------------------------------------------------------------
// Table List Window context menu items
//----------------------------------------------------------------------------
#define M_TLV_SELECT_ALL_FROM_TABLE             2601
#define M_TLV_BROWSE_TABLE                      2602
#define M_TLV_OPEN_NEW_MAP                      2603
#define M_TLV_EXPORT_TABLE                      2604
#define M_TLV_REFRESH_TABLE                     2605
#define M_TLV_TABLE_STRUCTURE                   2606
#define M_TLV_CREATE_POINTS                     2607
#define M_TLV_UPDATE_COLUMN                     2608
#define M_TLV_CLOSE_TABLE                       2609
#define M_TLV_COMPARE_STRUCTURE                 2612
#define M_TLV_RENAME_TABLE                      2613

//----------------------------------------------------------------------------
// Legend Designer menu items
//----------------------------------------------------------------------------
#define M_LEGEND_DESIGNER_HIDE_GUIDELINES       2801
#define M_LEGEND_DESIGNER_REMOVE_GUIDELINES     2802
#define M_LEGEND_DESIGNER_REMOVE_FRAMES         2803
#define M_LEGEND_DESIGNER_PROPERTIES            2804
#define M_LEGEND_DESIGNER_REFRESH               2805
#define M_LEGEND_DESIGNER_ADD_FRAMES            2806
#define M_LEGEND_DESIGNER_ADD_TEXTFRAME         2807
#define M_LEGEND_DESIGNER_MODIFY_THEME          2808
#define M_LEGEND_DESIGNER_TOGGLE_GRID           2809
#define M_LEGEND_DESIGNER_SNAP_TO_GRID          2810
#define M_LEGEND_DESIGNER_HORIZ_GUIDELINE       2811
#define M_LEGEND_DESIGNER_VERT_GUIDELINE        2812
#define M_LEGEND_DESIGNER_SNAP_TOP              2813
#define M_LEGEND_DESIGNER_SNAP_BOTTOM           2814
#define M_LEGEND_DESIGNER_SNAP_LEFT             2815
#define M_LEGEND_DESIGNER_SNAP_RIGHT            2816
#define M_LEGEND_DESIGNER_QUICK_REFRESH         2817

//----------------------------------------------------------------------------
// Menu and ButtonPad items that do not appear in the standard menus
//----------------------------------------------------------------------------
#define M_TOOLS_MAPBASIC                        1720
#define M_TOOLS_SEARCH_POLYGON                  1733

//----------------------------------------------------------------------------
// Codes for AdornmentInfo function to get info about an adornment win
//----------------------------------------------------------------------------
#define ADORNMENT_INFO_TYPE                            1
#define ADORNMENT_INFO_MAP_WINDOWID                    2
#define ADORNMENT_INFO_IS_FIXED_POS                    3
#define ADORNMENT_INFO_FIXED_POS_X                     4
#define ADORNMENT_INFO_FIXED_POS_Y                     5
#define ADORNMENT_INFO_FIXED_POS_UNITS                 6
#define ADORNMENT_INFO_DOCKED_POS                      7
#define ADORNMENT_INFO_DOCKED_OFFSET_X                 8
#define ADORNMENT_INFO_DOCKED_OFFSET_Y                 9
#define ADORNMENT_INFO_DOCKED_UNITS                   10
#define ADORNMENT_INFO_BACKGROUND_PEN                 11
#define ADORNMENT_INFO_BACKGROUND_BRUSH               12
#define ADORNMENT_INFO_SB_TYPE                        20
#define ADORNMENT_INFO_SB_MAP_UNITS                   21
#define ADORNMENT_INFO_SB_PAPER_UNITS                 22
#define ADORNMENT_INFO_SB_BAR_LENGTH                  23
#define ADORNMENT_INFO_SB_BAR_DRAW_LEN                24
#define ADORNMENT_INFO_SB_BAR_HEIGHT                  25
#define ADORNMENT_INFO_SB_AUTO_SCALING                26
#define ADORNMENT_INFO_SB_CARTO_SCALE                 27
#define ADORNMENT_INFO_SB_BAR_PEN                     28
#define ADORNMENT_INFO_SB_BAR_BRUSH                   29
#define ADORNMENT_INFO_SB_BAR_FONT                    30
#define ADORNMENT_INFO_SB_DISPLAY_SCALE               31
#define ADORNMENT_INFO_SB_AUTOOFF_SCALE               32
#define ADORNMENT_INFO_SB_AUTOON_SCALE                33
#define ADORNMENT_INFO_SB_SCALE_STRING                34
#define ADORNMENT_INFO_SB_CARTO_VALUE                 35
#define ADORNMENT_INFO_SB_CARTO_STRING                36

//----------------------------------------------------------------------------
// Codes used to position Adornments relative to mapper
//----------------------------------------------------------------------------

#define ADORNMENT_INFO_MAP_POS_TL     0
#define ADORNMENT_INFO_MAP_POS_TC     1
#define ADORNMENT_INFO_MAP_POS_TR     2
#define ADORNMENT_INFO_MAP_POS_CL     3
#define ADORNMENT_INFO_MAP_POS_CC     4
#define ADORNMENT_INFO_MAP_POS_CR     5
#define ADORNMENT_INFO_MAP_POS_BL     6
#define ADORNMENT_INFO_MAP_POS_BC     7
#define ADORNMENT_INFO_MAP_POS_BR     8
#define SCALEBAR_INFO_BARTYPE_CHECKEDBAR  0
#define SCALEBAR_INFO_BARTYPE_SOLIDBAR    1
#define SCALEBAR_INFO_BARTYPE_LINEBAR     2
#define SCALEBAR_INFO_BARTYPE_TICKBAR     3

//----------------------------------------------------------------------------
// Assiging nullptr to c# ref type will destory them and allow GC to collect
//----------------------------------------------------------------------------
#define NULL_PTR                         -1

//----------------------------------------------------------------------------
// Coordinate system datum id's.  These match the id's from mapinfow.prj.
//----------------------------------------------------------------------------
#define DATUMID_NAD27  62
#define DATUMID_NAD83  74
#define DATUMID_WGS84 104

//----------------------------------------------------------------------------
// Menu defines - these represent the predefined menus like File, and can
// be used in the Alter Menu command.  These defines cannot be used in a
// "Run Menu Command" statement.
//----------------------------------------------------------------------------
#define M_FILE                                    1
#define M_EDIT                                    2
#define M_SEARCH                                  3
#define M_QUERY                                   3
#define M_PGM                                     4
#define M_TOOLS                                   4
#define M_OPTIONS                                 5
#define M_WINDOW                                  6
#define M_HELP                                    7
#define M_BROWSE                                  8
#define M_MAP                                     9
#define M_LAYOUT                                 10
#define M_MAPBASIC                               12
#define M_REDISTRICT                             13
#define M_OBJECTS                                14
#define M_TABLE                                  15
#define M_SHORTCUT_DFLT                          16
#define M_SHORTCUT_MAPPER                        17
#define M_SHORTCUT_BROWSER                       18
#define M_SHORTCUT_LAYOUT                        19
#define M_SHORTCUT_CMD                           21
#define M_SHORTCUT_REDISTRICTER                  22
#define M_SHORTCUT_LEGEND                        23
#define M_SHORTCUT_GRAPHTDG                      24
#define M_SHORTCUT_3DMAP                         25
#define M_SHORTCUT_MSG_WIN                       26
#define M_SHORTCUT_STAT_WIN                      27
#define M_3DMAP                                  28
#define M_GRAPHTDG                               29
#define M_LEGEND                                 31
#define M_SHORTCUT_ADORNMENT                     32
#define M_SHORTCUT_LC_LAYERS                     33
#define M_SHORTCUT_LC_MAPS                       34
#define M_SHORTCUT_LC_GROUPS                     35
#define M_SHORTCUT_TABLEADORNMENT                36
#define M_SHORTCUT_TLV_TABLES                    37
#define M_SHORTCUT_LC_OVERRIDES                  38
#define M_OPEN_WEBSERVICE                        39
#define M_FIND_SELECTION                         40
#define M_ORACLE_TOOLS                           41
#define M_TAB_MAINTENANCE                        42
#define M_TAB_RASTER                             43
#define M_LICENSING                              44
#define M_TILESERVER                             45
#define M_SHORTCUT_BROWSER_SORTFILTER            47
#define M_SHORTCUT_BROWSER_SELECT                48
#define M_LEGEND_DESIGNER                        49
#define M_SHORTCUT_LEGEND_DESIGNER               50
#define M_LAYOUT_DESIGNER                        51

//----------------------------------------------------------------------------
// Buttonpad defines - these represent the predefined toolbars, and can
// be used in the Alter ButtonPad command.
//----------------------------------------------------------------------------
#define TB_MAIN                                   1
#define TB_DRAWING                                2
#define TB_TOOLS                                  3
#define TB_STANDARD                               4
#define TB_ODBC                                   5
#define TB_WEB_SVCS                               6

//============================================================================
// end of MENU.DEF
//============================================================================
