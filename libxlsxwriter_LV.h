/*
 * libxlsxwriter_LV.h - Combined public API header for libxlsxwriter
 *
 * This header contains all public exported API functions with standard
 * C calling convention (__stdcall/WINAPI) declarations.
 * All macros have been expanded to their equivalent code.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 *
 * Generated for LabVIEW integration.
 */

#ifndef __LIBXLSXWRITER_LV_H__
#define __LIBXLSXWRITER_LV_H__

#include <stdint.h>
#include <stdio.h>
#include <time.h>

#ifdef __cplusplus
extern "C" {
#endif

/* ============================================================================
 * Calling Convention Definition
 * ============================================================================ */

#if defined(_WIN32) || defined(_WIN64)
    #define LXW_CALL __stdcall
#else
    #define LXW_CALL
#endif

/* ============================================================================
 * Basic Type Definitions
 * ============================================================================ */

/** Integer data type to represent a row value (max 1,048,576). */
typedef uint32_t lxw_row_t;

/** Integer data type to represent a column value (max 16,384). */
typedef uint16_t lxw_col_t;

/** The type for RGB colors (0x000000 to 0xFFFFFF). */
typedef uint32_t lxw_color_t;

/* ============================================================================
 * Boolean Values
 * ============================================================================ */

enum lxw_boolean {
    LXW_FALSE = 0,
    LXW_TRUE = 1,
    LXW_EXPLICIT_FALSE = 2
};

/* ============================================================================
 * Error Codes
 * ============================================================================ */

typedef enum lxw_error {
    LXW_NO_ERROR = 0,
    LXW_ERROR_MEMORY_MALLOC_FAILED,
    LXW_ERROR_CREATING_XLSX_FILE,
    LXW_ERROR_CREATING_TMPFILE,
    LXW_ERROR_READING_TMPFILE,
    LXW_ERROR_ZIP_FILE_OPERATION,
    LXW_ERROR_ZIP_PARAMETER_ERROR,
    LXW_ERROR_ZIP_BAD_ZIP_FILE,
    LXW_ERROR_ZIP_INTERNAL_ERROR,
    LXW_ERROR_ZIP_FILE_ADD,
    LXW_ERROR_ZIP_CLOSE,
    LXW_ERROR_FEATURE_NOT_SUPPORTED,
    LXW_ERROR_NULL_PARAMETER_IGNORED,
    LXW_ERROR_PARAMETER_VALIDATION,
    LXW_ERROR_PARAMETER_IS_EMPTY,
    LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED,
    LXW_ERROR_INVALID_SHEETNAME_CHARACTER,
    LXW_ERROR_SHEETNAME_START_END_APOSTROPHE,
    LXW_ERROR_SHEETNAME_ALREADY_USED,
    LXW_ERROR_32_STRING_LENGTH_EXCEEDED,
    LXW_ERROR_128_STRING_LENGTH_EXCEEDED,
    LXW_ERROR_255_STRING_LENGTH_EXCEEDED,
    LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED,
    LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND,
    LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE,
    LXW_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED,
    LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED,
    LXW_ERROR_IMAGE_DIMENSIONS,
    LXW_MAX_ERRNO
} lxw_error;

/* ============================================================================
 * Datetime Structure
 * ============================================================================ */

typedef struct lxw_datetime {
    int year;       /* 1900 - 9999 */
    int month;      /* 1 - 12 */
    int day;        /* 1 - 31 */
    int hour;       /* 0 - 23 */
    int min;        /* 0 - 59 */
    double sec;     /* 0 - 59.999 */
} lxw_datetime;

/* ============================================================================
 * Chart Types
 * ============================================================================ */

typedef enum lxw_chart_type {
    LXW_CHART_NONE = 0,
    LXW_CHART_AREA,
    LXW_CHART_AREA_STACKED,
    LXW_CHART_AREA_STACKED_PERCENT,
    LXW_CHART_BAR,
    LXW_CHART_BAR_STACKED,
    LXW_CHART_BAR_STACKED_PERCENT,
    LXW_CHART_COLUMN,
    LXW_CHART_COLUMN_STACKED,
    LXW_CHART_COLUMN_STACKED_PERCENT,
    LXW_CHART_DOUGHNUT,
    LXW_CHART_LINE,
    LXW_CHART_LINE_STACKED,
    LXW_CHART_LINE_STACKED_PERCENT,
    LXW_CHART_PIE,
    LXW_CHART_SCATTER,
    LXW_CHART_SCATTER_STRAIGHT,
    LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS,
    LXW_CHART_SCATTER_SMOOTH,
    LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS,
    LXW_CHART_RADAR,
    LXW_CHART_RADAR_WITH_MARKERS,
    LXW_CHART_RADAR_FILLED
} lxw_chart_type;

typedef enum lxw_chart_legend_position {
    LXW_CHART_LEGEND_NONE = 0,
    LXW_CHART_LEGEND_RIGHT,
    LXW_CHART_LEGEND_LEFT,
    LXW_CHART_LEGEND_TOP,
    LXW_CHART_LEGEND_BOTTOM,
    LXW_CHART_LEGEND_TOP_RIGHT,
    LXW_CHART_LEGEND_OVERLAY_RIGHT,
    LXW_CHART_LEGEND_OVERLAY_LEFT,
    LXW_CHART_LEGEND_OVERLAY_TOP_RIGHT
} lxw_chart_legend_position;

typedef enum lxw_chart_line_dash_type {
    LXW_CHART_LINE_DASH_SOLID = 0,
    LXW_CHART_LINE_DASH_ROUND_DOT,
    LXW_CHART_LINE_DASH_SQUARE_DOT,
    LXW_CHART_LINE_DASH_DASH,
    LXW_CHART_LINE_DASH_DASH_DOT,
    LXW_CHART_LINE_DASH_LONG_DASH,
    LXW_CHART_LINE_DASH_LONG_DASH_DOT,
    LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT,
    LXW_CHART_LINE_DASH_DOT,
    LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT,
    LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT
} lxw_chart_line_dash_type;

typedef enum lxw_chart_marker_type {
    LXW_CHART_MARKER_AUTOMATIC,
    LXW_CHART_MARKER_NONE,
    LXW_CHART_MARKER_SQUARE,
    LXW_CHART_MARKER_DIAMOND,
    LXW_CHART_MARKER_TRIANGLE,
    LXW_CHART_MARKER_X,
    LXW_CHART_MARKER_STAR,
    LXW_CHART_MARKER_SHORT_DASH,
    LXW_CHART_MARKER_LONG_DASH,
    LXW_CHART_MARKER_CIRCLE,
    LXW_CHART_MARKER_PLUS
} lxw_chart_marker_type;

typedef enum lxw_chart_label_position {
    LXW_CHART_LABEL_POSITION_DEFAULT,
    LXW_CHART_LABEL_POSITION_CENTER,
    LXW_CHART_LABEL_POSITION_RIGHT,
    LXW_CHART_LABEL_POSITION_LEFT,
    LXW_CHART_LABEL_POSITION_ABOVE,
    LXW_CHART_LABEL_POSITION_BELOW,
    LXW_CHART_LABEL_POSITION_INSIDE_BASE,
    LXW_CHART_LABEL_POSITION_INSIDE_END,
    LXW_CHART_LABEL_POSITION_OUTSIDE_END,
    LXW_CHART_LABEL_POSITION_BEST_FIT
} lxw_chart_label_position;

typedef enum lxw_chart_label_separator {
    LXW_CHART_LABEL_SEPARATOR_COMMA,
    LXW_CHART_LABEL_SEPARATOR_SEMICOLON,
    LXW_CHART_LABEL_SEPARATOR_PERIOD,
    LXW_CHART_LABEL_SEPARATOR_NEWLINE,
    LXW_CHART_LABEL_SEPARATOR_SPACE
} lxw_chart_label_separator;

typedef enum lxw_chart_axis_tick_position {
    LXW_CHART_AXIS_POSITION_DEFAULT,
    LXW_CHART_AXIS_POSITION_ON_TICK,
    LXW_CHART_AXIS_POSITION_BETWEEN
} lxw_chart_axis_tick_position;

typedef enum lxw_chart_axis_label_position {
    LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO,
    LXW_CHART_AXIS_LABEL_POSITION_HIGH,
    LXW_CHART_AXIS_LABEL_POSITION_LOW,
    LXW_CHART_AXIS_LABEL_POSITION_NONE
} lxw_chart_axis_label_position;

typedef enum lxw_chart_axis_label_alignment {
    LXW_CHART_AXIS_LABEL_ALIGN_CENTER,
    LXW_CHART_AXIS_LABEL_ALIGN_LEFT,
    LXW_CHART_AXIS_LABEL_ALIGN_RIGHT
} lxw_chart_axis_label_alignment;

typedef enum lxw_chart_axis_display_unit {
    LXW_CHART_AXIS_UNITS_NONE,
    LXW_CHART_AXIS_UNITS_HUNDREDS,
    LXW_CHART_AXIS_UNITS_THOUSANDS,
    LXW_CHART_AXIS_UNITS_TEN_THOUSANDS,
    LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS,
    LXW_CHART_AXIS_UNITS_MILLIONS,
    LXW_CHART_AXIS_UNITS_TEN_MILLIONS,
    LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS,
    LXW_CHART_AXIS_UNITS_BILLIONS,
    LXW_CHART_AXIS_UNITS_TRILLIONS
} lxw_chart_axis_display_unit;

typedef enum lxw_chart_axis_tick_mark {
    LXW_CHART_AXIS_TICK_MARK_DEFAULT,
    LXW_CHART_AXIS_TICK_MARK_NONE,
    LXW_CHART_AXIS_TICK_MARK_INSIDE,
    LXW_CHART_AXIS_TICK_MARK_OUTSIDE,
    LXW_CHART_AXIS_TICK_MARK_CROSSING
} lxw_chart_tick_mark;

typedef enum lxw_chart_blank {
    LXW_CHART_BLANKS_AS_GAP,
    LXW_CHART_BLANKS_AS_ZERO,
    LXW_CHART_BLANKS_AS_CONNECTED
} lxw_chart_blank;

typedef enum lxw_chart_error_bar_type {
    LXW_CHART_ERROR_BAR_TYPE_STD_ERROR,
    LXW_CHART_ERROR_BAR_TYPE_FIXED,
    LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE,
    LXW_CHART_ERROR_BAR_TYPE_STD_DEV
} lxw_chart_error_bar_type;

typedef enum lxw_chart_error_bar_direction {
    LXW_CHART_ERROR_BAR_DIR_BOTH,
    LXW_CHART_ERROR_BAR_DIR_PLUS,
    LXW_CHART_ERROR_BAR_DIR_MINUS
} lxw_chart_error_bar_direction;

typedef enum lxw_chart_error_bar_cap {
    LXW_CHART_ERROR_BAR_END_CAP,
    LXW_CHART_ERROR_BAR_NO_CAP
} lxw_chart_error_bar_cap;

typedef enum lxw_chart_trendline_type {
    LXW_CHART_TRENDLINE_TYPE_LINEAR,
    LXW_CHART_TRENDLINE_TYPE_LOG,
    LXW_CHART_TRENDLINE_TYPE_POLY,
    LXW_CHART_TRENDLINE_TYPE_POWER,
    LXW_CHART_TRENDLINE_TYPE_EXP,
    LXW_CHART_TRENDLINE_TYPE_AVERAGE
} lxw_chart_trendline_type;

/* ============================================================================
 * Format Enums
 * ============================================================================ */

enum lxw_format_underlines {
    LXW_UNDERLINE_NONE = 0,
    LXW_UNDERLINE_SINGLE,
    LXW_UNDERLINE_DOUBLE,
    LXW_UNDERLINE_SINGLE_ACCOUNTING,
    LXW_UNDERLINE_DOUBLE_ACCOUNTING
};

enum lxw_format_scripts {
    LXW_FONT_SUPERSCRIPT = 1,
    LXW_FONT_SUBSCRIPT
};

enum lxw_format_alignments {
    LXW_ALIGN_NONE = 0,
    LXW_ALIGN_LEFT,
    LXW_ALIGN_CENTER,
    LXW_ALIGN_RIGHT,
    LXW_ALIGN_FILL,
    LXW_ALIGN_JUSTIFY,
    LXW_ALIGN_CENTER_ACROSS,
    LXW_ALIGN_DISTRIBUTED,
    LXW_ALIGN_VERTICAL_TOP,
    LXW_ALIGN_VERTICAL_BOTTOM,
    LXW_ALIGN_VERTICAL_CENTER,
    LXW_ALIGN_VERTICAL_JUSTIFY,
    LXW_ALIGN_VERTICAL_DISTRIBUTED
};

enum lxw_format_diagonal_types {
    LXW_DIAGONAL_BORDER_UP = 1,
    LXW_DIAGONAL_BORDER_DOWN,
    LXW_DIAGONAL_BORDER_UP_DOWN
};

enum lxw_defined_colors {
    LXW_COLOR_BLACK = 0x1000000,
    LXW_COLOR_BLUE = 0x0000FF,
    LXW_COLOR_BROWN = 0x800000,
    LXW_COLOR_CYAN = 0x00FFFF,
    LXW_COLOR_GRAY = 0x808080,
    LXW_COLOR_GREEN = 0x008000,
    LXW_COLOR_LIME = 0x00FF00,
    LXW_COLOR_MAGENTA = 0xFF00FF,
    LXW_COLOR_NAVY = 0x000080,
    LXW_COLOR_ORANGE = 0xFF6600,
    LXW_COLOR_PINK = 0xFF00FF,
    LXW_COLOR_PURPLE = 0x800080,
    LXW_COLOR_RED = 0xFF0000,
    LXW_COLOR_SILVER = 0xC0C0C0,
    LXW_COLOR_WHITE = 0xFFFFFF,
    LXW_COLOR_YELLOW = 0xFFFF00
};

enum lxw_format_patterns {
    LXW_PATTERN_NONE = 0,
    LXW_PATTERN_SOLID,
    LXW_PATTERN_MEDIUM_GRAY,
    LXW_PATTERN_DARK_GRAY,
    LXW_PATTERN_LIGHT_GRAY,
    LXW_PATTERN_DARK_HORIZONTAL,
    LXW_PATTERN_DARK_VERTICAL,
    LXW_PATTERN_DARK_DOWN,
    LXW_PATTERN_DARK_UP,
    LXW_PATTERN_DARK_GRID,
    LXW_PATTERN_DARK_TRELLIS,
    LXW_PATTERN_LIGHT_HORIZONTAL,
    LXW_PATTERN_LIGHT_VERTICAL,
    LXW_PATTERN_LIGHT_DOWN,
    LXW_PATTERN_LIGHT_UP,
    LXW_PATTERN_LIGHT_GRID,
    LXW_PATTERN_LIGHT_TRELLIS,
    LXW_PATTERN_GRAY_125,
    LXW_PATTERN_GRAY_0625
};

enum lxw_format_borders {
    LXW_BORDER_NONE,
    LXW_BORDER_THIN,
    LXW_BORDER_MEDIUM,
    LXW_BORDER_DASHED,
    LXW_BORDER_DOTTED,
    LXW_BORDER_THICK,
    LXW_BORDER_DOUBLE,
    LXW_BORDER_HAIR,
    LXW_BORDER_MEDIUM_DASHED,
    LXW_BORDER_DASH_DOT,
    LXW_BORDER_MEDIUM_DASH_DOT,
    LXW_BORDER_DASH_DOT_DOT,
    LXW_BORDER_MEDIUM_DASH_DOT_DOT,
    LXW_BORDER_SLANT_DASH_DOT
};

/* ============================================================================
 * Forward Declarations (Opaque Pointers)
 * ============================================================================ */

typedef struct lxw_workbook lxw_workbook;
typedef struct lxw_worksheet lxw_worksheet;
typedef struct lxw_chartsheet lxw_chartsheet;
typedef struct lxw_chart lxw_chart;
typedef struct lxw_chart_series lxw_chart_series;
typedef struct lxw_chart_axis lxw_chart_axis;
typedef struct lxw_format lxw_format;
typedef struct lxw_chart_line lxw_chart_line;
typedef struct lxw_chart_fill lxw_chart_fill;
typedef struct lxw_chart_pattern lxw_chart_pattern;
typedef struct lxw_chart_font lxw_chart_font;
typedef struct lxw_chart_layout lxw_chart_layout;
typedef struct lxw_chart_point lxw_chart_point;
typedef struct lxw_chart_data_label lxw_chart_data_label;
typedef struct lxw_series_error_bars lxw_series_error_bars;
typedef struct lxw_row_col_options lxw_row_col_options;
typedef struct lxw_image_options lxw_image_options;
typedef struct lxw_chart_options lxw_chart_options;
typedef struct lxw_protection lxw_protection;
typedef struct lxw_header_footer_options lxw_header_footer_options;
typedef struct lxw_data_validation lxw_data_validation;
typedef struct lxw_conditional_format lxw_conditional_format;
typedef struct lxw_button_options lxw_button_options;
typedef struct lxw_table_options lxw_table_options;
typedef struct lxw_filter_rule lxw_filter_rule;
typedef struct lxw_rich_string_tuple lxw_rich_string_tuple;
typedef struct lxw_comment_options lxw_comment_options;
typedef struct lxw_workbook_options lxw_workbook_options;
typedef struct lxw_doc_properties lxw_doc_properties;

/* ============================================================================
 * Chart Line/Fill/Pattern Structures (for inline initialization)
 * ============================================================================ */

struct lxw_chart_line {
    lxw_color_t color;
    uint8_t none;
    float width;
    uint8_t dash_type;
    uint8_t transparency;
};

struct lxw_chart_fill {
    lxw_color_t color;
    uint8_t none;
    uint8_t transparency;
};

struct lxw_chart_pattern {
    lxw_color_t fg_color;
    lxw_color_t bg_color;
    uint8_t type;
};

struct lxw_chart_font {
    const char *name;
    double size;
    uint8_t bold;
    uint8_t italic;
    uint8_t underline;
    int32_t rotation;
    lxw_color_t color;
    uint8_t pitch_family;
    uint8_t charset;
    int8_t baseline;
};

struct lxw_chart_layout {
    double x;
    double y;
    double width;
    double height;
    uint8_t has_inner;
};

/* ============================================================================
 * WORKBOOK FUNCTIONS
 * ============================================================================ */

lxw_workbook * LXW_CALL workbook_new(const char *filename);
lxw_workbook * LXW_CALL workbook_new_opt(const char *filename, lxw_workbook_options *options);
lxw_worksheet * LXW_CALL workbook_add_worksheet(lxw_workbook *workbook, const char *sheetname);
lxw_chartsheet * LXW_CALL workbook_add_chartsheet(lxw_workbook *workbook, const char *sheetname);
lxw_format * LXW_CALL workbook_add_format(lxw_workbook *workbook);
lxw_chart * LXW_CALL workbook_add_chart(lxw_workbook *workbook, uint8_t chart_type);
lxw_error LXW_CALL workbook_close(lxw_workbook *workbook);
lxw_error LXW_CALL workbook_set_properties(lxw_workbook *workbook, lxw_doc_properties *properties);
lxw_error LXW_CALL workbook_set_custom_property_string(lxw_workbook *workbook, const char *name, const char *value);
lxw_error LXW_CALL workbook_set_custom_property_number(lxw_workbook *workbook, const char *name, double value);
lxw_error LXW_CALL workbook_set_custom_property_integer(lxw_workbook *workbook, const char *name, int32_t value);
lxw_error LXW_CALL workbook_set_custom_property_boolean(lxw_workbook *workbook, const char *name, uint8_t value);
lxw_error LXW_CALL workbook_set_custom_property_datetime(lxw_workbook *workbook, const char *name, lxw_datetime *datetime);
lxw_error LXW_CALL workbook_define_name(lxw_workbook *workbook, const char *name, const char *formula);
lxw_format * LXW_CALL workbook_get_default_url_format(lxw_workbook *workbook);
lxw_worksheet * LXW_CALL workbook_get_worksheet_by_name(lxw_workbook *workbook, const char *name);
lxw_chartsheet * LXW_CALL workbook_get_chartsheet_by_name(lxw_workbook *workbook, const char *name);
lxw_error LXW_CALL workbook_validate_sheet_name(lxw_workbook *workbook, const char *sheetname);
lxw_error LXW_CALL workbook_add_vba_project(lxw_workbook *workbook, const char *filename);
lxw_error LXW_CALL workbook_add_signed_vba_project(lxw_workbook *workbook, const char *vba_project, const char *signature);
lxw_error LXW_CALL workbook_set_vba_name(lxw_workbook *workbook, const char *name);
void LXW_CALL workbook_read_only_recommended(lxw_workbook *workbook);
void LXW_CALL workbook_use_1904_epoch(lxw_workbook *workbook);
void LXW_CALL workbook_set_size(lxw_workbook *workbook, uint16_t width, uint16_t height);

/* ============================================================================
 * WORKSHEET FUNCTIONS
 * ============================================================================ */

lxw_error LXW_CALL worksheet_write_number(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, double number, lxw_format *format);
lxw_error LXW_CALL worksheet_write_string(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *string, lxw_format *format);
lxw_error LXW_CALL worksheet_write_formula(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format *format);
lxw_error LXW_CALL worksheet_write_array_formula(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format *format);
lxw_error LXW_CALL worksheet_write_dynamic_array_formula(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format *format);
lxw_error LXW_CALL worksheet_write_dynamic_formula(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format *format);
lxw_error LXW_CALL worksheet_write_datetime(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_datetime *datetime, lxw_format *format);
lxw_error LXW_CALL worksheet_write_unixtime(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, int64_t unixtime, lxw_format *format);
lxw_error LXW_CALL worksheet_write_url(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *url, lxw_format *format);
lxw_error LXW_CALL worksheet_write_url_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *url, lxw_format *format, const char *string, const char *tooltip);
lxw_error LXW_CALL worksheet_write_boolean(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, int value, lxw_format *format);
lxw_error LXW_CALL worksheet_write_blank(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_format *format);
lxw_error LXW_CALL worksheet_write_formula_num(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format *format, double result);
lxw_error LXW_CALL worksheet_write_formula_str(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format *format, const char *result);
lxw_error LXW_CALL worksheet_write_rich_string(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_rich_string_tuple *rich_string[], lxw_format *format);
lxw_error LXW_CALL worksheet_write_comment(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *string);
lxw_error LXW_CALL worksheet_write_comment_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *string, lxw_comment_options *options);

lxw_error LXW_CALL worksheet_set_row(lxw_worksheet *worksheet, lxw_row_t row, double height, lxw_format *format);
lxw_error LXW_CALL worksheet_set_row_opt(lxw_worksheet *worksheet, lxw_row_t row, double height, lxw_format *format, lxw_row_col_options *options);
lxw_error LXW_CALL worksheet_set_row_pixels(lxw_worksheet *worksheet, lxw_row_t row, uint32_t pixels, lxw_format *format);
lxw_error LXW_CALL worksheet_set_row_pixels_opt(lxw_worksheet *worksheet, lxw_row_t row, uint32_t pixels, lxw_format *format, lxw_row_col_options *options);
lxw_error LXW_CALL worksheet_set_column(lxw_worksheet *worksheet, lxw_col_t first_col, lxw_col_t last_col, double width, lxw_format *format);
lxw_error LXW_CALL worksheet_set_column_opt(lxw_worksheet *worksheet, lxw_col_t first_col, lxw_col_t last_col, double width, lxw_format *format, lxw_row_col_options *options);
lxw_error LXW_CALL worksheet_set_column_pixels(lxw_worksheet *worksheet, lxw_col_t first_col, lxw_col_t last_col, uint32_t pixels, lxw_format *format);
lxw_error LXW_CALL worksheet_set_column_pixels_opt(lxw_worksheet *worksheet, lxw_col_t first_col, lxw_col_t last_col, uint32_t pixels, lxw_format *format, lxw_row_col_options *options);

lxw_error LXW_CALL worksheet_insert_image(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *filename);
lxw_error LXW_CALL worksheet_insert_image_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *filename, lxw_image_options *options);
lxw_error LXW_CALL worksheet_insert_image_buffer(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size);
lxw_error LXW_CALL worksheet_insert_image_buffer_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size, lxw_image_options *options);
lxw_error LXW_CALL worksheet_embed_image(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *filename);
lxw_error LXW_CALL worksheet_embed_image_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const char *filename, lxw_image_options *options);
lxw_error LXW_CALL worksheet_embed_image_buffer(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size);
lxw_error LXW_CALL worksheet_embed_image_buffer_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size, lxw_image_options *options);
lxw_error LXW_CALL worksheet_set_background(lxw_worksheet *worksheet, const char *filename);
lxw_error LXW_CALL worksheet_set_background_buffer(lxw_worksheet *worksheet, const unsigned char *image_buffer, size_t image_size);

lxw_error LXW_CALL worksheet_insert_chart(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_chart *chart);
lxw_error LXW_CALL worksheet_insert_chart_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_chart *chart, lxw_chart_options *options);
lxw_error LXW_CALL worksheet_merge_range(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *string, lxw_format *format);
lxw_error LXW_CALL worksheet_autofilter(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
lxw_error LXW_CALL worksheet_filter_column(lxw_worksheet *worksheet, lxw_col_t col, lxw_filter_rule *rule);
lxw_error LXW_CALL worksheet_filter_column2(lxw_worksheet *worksheet, lxw_col_t col, lxw_filter_rule *rule1, lxw_filter_rule *rule2, uint8_t and_or);
lxw_error LXW_CALL worksheet_filter_list(lxw_worksheet *worksheet, lxw_col_t col, const char **list);
lxw_error LXW_CALL worksheet_data_validation_cell(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_data_validation *validation);
lxw_error LXW_CALL worksheet_data_validation_range(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, lxw_data_validation *validation);
lxw_error LXW_CALL worksheet_conditional_format_cell(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_conditional_format *conditional_format);
lxw_error LXW_CALL worksheet_conditional_format_range(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, lxw_conditional_format *conditional_format);
lxw_error LXW_CALL worksheet_insert_button(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_button_options *options);
lxw_error LXW_CALL worksheet_add_table(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, lxw_table_options *options);

void LXW_CALL worksheet_activate(lxw_worksheet *worksheet);
void LXW_CALL worksheet_select(lxw_worksheet *worksheet);
void LXW_CALL worksheet_hide(lxw_worksheet *worksheet);
void LXW_CALL worksheet_set_first_sheet(lxw_worksheet *worksheet);
void LXW_CALL worksheet_freeze_panes(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col);
void LXW_CALL worksheet_split_panes(lxw_worksheet *worksheet, double vertical, double horizontal);
lxw_error LXW_CALL worksheet_set_selection(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void LXW_CALL worksheet_set_top_left_cell(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col);
void LXW_CALL worksheet_set_landscape(lxw_worksheet *worksheet);
void LXW_CALL worksheet_set_portrait(lxw_worksheet *worksheet);
void LXW_CALL worksheet_set_page_view(lxw_worksheet *worksheet);
void LXW_CALL worksheet_set_paper(lxw_worksheet *worksheet, uint8_t paper_type);
void LXW_CALL worksheet_set_margins(lxw_worksheet *worksheet, double left, double right, double top, double bottom);
lxw_error LXW_CALL worksheet_set_header(lxw_worksheet *worksheet, const char *string);
lxw_error LXW_CALL worksheet_set_footer(lxw_worksheet *worksheet, const char *string);
lxw_error LXW_CALL worksheet_set_header_opt(lxw_worksheet *worksheet, const char *string, lxw_header_footer_options *options);
lxw_error LXW_CALL worksheet_set_footer_opt(lxw_worksheet *worksheet, const char *string, lxw_header_footer_options *options);
lxw_error LXW_CALL worksheet_set_h_pagebreaks(lxw_worksheet *worksheet, lxw_row_t breaks[]);
lxw_error LXW_CALL worksheet_set_v_pagebreaks(lxw_worksheet *worksheet, lxw_col_t breaks[]);
void LXW_CALL worksheet_print_across(lxw_worksheet *worksheet);
void LXW_CALL worksheet_set_zoom(lxw_worksheet *worksheet, uint16_t scale);
void LXW_CALL worksheet_gridlines(lxw_worksheet *worksheet, uint8_t option);
void LXW_CALL worksheet_center_horizontally(lxw_worksheet *worksheet);
void LXW_CALL worksheet_center_vertically(lxw_worksheet *worksheet);
void LXW_CALL worksheet_print_row_col_headers(lxw_worksheet *worksheet);
lxw_error LXW_CALL worksheet_repeat_rows(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_row_t last_row);
lxw_error LXW_CALL worksheet_repeat_columns(lxw_worksheet *worksheet, lxw_col_t first_col, lxw_col_t last_col);
lxw_error LXW_CALL worksheet_print_area(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void LXW_CALL worksheet_fit_to_pages(lxw_worksheet *worksheet, uint16_t width, uint16_t height);
void LXW_CALL worksheet_set_start_page(lxw_worksheet *worksheet, uint16_t start_page);
void LXW_CALL worksheet_set_print_scale(lxw_worksheet *worksheet, uint16_t scale);
void LXW_CALL worksheet_print_black_and_white(lxw_worksheet *worksheet);
void LXW_CALL worksheet_right_to_left(lxw_worksheet *worksheet);
void LXW_CALL worksheet_hide_zero(lxw_worksheet *worksheet);
void LXW_CALL worksheet_set_tab_color(lxw_worksheet *worksheet, lxw_color_t color);
void LXW_CALL worksheet_protect(lxw_worksheet *worksheet, const char *password, lxw_protection *options);
void LXW_CALL worksheet_outline_settings(lxw_worksheet *worksheet, uint8_t visible, uint8_t symbols_below, uint8_t symbols_right, uint8_t auto_style);
void LXW_CALL worksheet_set_default_row(lxw_worksheet *worksheet, double height, uint8_t hide_unused_rows);
lxw_error LXW_CALL worksheet_set_vba_name(lxw_worksheet *worksheet, const char *name);
void LXW_CALL worksheet_show_comments(lxw_worksheet *worksheet);
void LXW_CALL worksheet_set_comments_author(lxw_worksheet *worksheet, const char *author);
lxw_error LXW_CALL worksheet_ignore_errors(lxw_worksheet *worksheet, uint8_t type, const char *range);

/* ============================================================================
 * CHARTSHEET FUNCTIONS
 * ============================================================================ */

lxw_error LXW_CALL chartsheet_set_chart(lxw_chartsheet *chartsheet, lxw_chart *chart);
lxw_error LXW_CALL chartsheet_set_chart_opt(lxw_chartsheet *chartsheet, lxw_chart *chart, lxw_chart_options *user_options);
void LXW_CALL chartsheet_activate(lxw_chartsheet *chartsheet);
void LXW_CALL chartsheet_select(lxw_chartsheet *chartsheet);
void LXW_CALL chartsheet_hide(lxw_chartsheet *chartsheet);
void LXW_CALL chartsheet_set_first_sheet(lxw_chartsheet *chartsheet);
void LXW_CALL chartsheet_set_tab_color(lxw_chartsheet *chartsheet, lxw_color_t color);
void LXW_CALL chartsheet_protect(lxw_chartsheet *chartsheet, const char *password, lxw_protection *options);
void LXW_CALL chartsheet_set_zoom(lxw_chartsheet *chartsheet, uint16_t scale);
void LXW_CALL chartsheet_set_landscape(lxw_chartsheet *chartsheet);
void LXW_CALL chartsheet_set_portrait(lxw_chartsheet *chartsheet);
void LXW_CALL chartsheet_set_paper(lxw_chartsheet *chartsheet, uint8_t paper_type);
void LXW_CALL chartsheet_set_margins(lxw_chartsheet *chartsheet, double left, double right, double top, double bottom);
lxw_error LXW_CALL chartsheet_set_header(lxw_chartsheet *chartsheet, const char *string);
lxw_error LXW_CALL chartsheet_set_footer(lxw_chartsheet *chartsheet, const char *string);
lxw_error LXW_CALL chartsheet_set_header_opt(lxw_chartsheet *chartsheet, const char *string, lxw_header_footer_options *options);
lxw_error LXW_CALL chartsheet_set_footer_opt(lxw_chartsheet *chartsheet, const char *string, lxw_header_footer_options *options);

/* ============================================================================
 * CHART FUNCTIONS
 * ============================================================================ */

lxw_chart_series * LXW_CALL chart_add_series(lxw_chart *chart, const char *categories, const char *values);
lxw_chart_series * LXW_CALL chart_add_series_on_axis(lxw_chart *chart, const char *categories, const char *values, lxw_chart_axis *y_axis);

void LXW_CALL chart_series_set_categories(lxw_chart_series *series, const char *sheetname, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void LXW_CALL chart_series_set_values(lxw_chart_series *series, const char *sheetname, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void LXW_CALL chart_series_set_name(lxw_chart_series *series, const char *name);
void LXW_CALL chart_series_set_name_range(lxw_chart_series *series, const char *sheetname, lxw_row_t row, lxw_col_t col);
void LXW_CALL chart_series_set_line(lxw_chart_series *series, lxw_chart_line *line);
void LXW_CALL chart_series_set_fill(lxw_chart_series *series, lxw_chart_fill *fill);
void LXW_CALL chart_series_set_invert_if_negative(lxw_chart_series *series);
void LXW_CALL chart_series_set_pattern(lxw_chart_series *series, lxw_chart_pattern *pattern);
void LXW_CALL chart_series_set_marker_type(lxw_chart_series *series, uint8_t type);
void LXW_CALL chart_series_set_marker_size(lxw_chart_series *series, uint8_t size);
void LXW_CALL chart_series_set_marker_line(lxw_chart_series *series, lxw_chart_line *line);
void LXW_CALL chart_series_set_marker_fill(lxw_chart_series *series, lxw_chart_fill *fill);
void LXW_CALL chart_series_set_marker_pattern(lxw_chart_series *series, lxw_chart_pattern *pattern);
lxw_error LXW_CALL chart_series_set_points(lxw_chart_series *series, lxw_chart_point *points[]);
void LXW_CALL chart_series_set_smooth(lxw_chart_series *series, uint8_t smooth);
void LXW_CALL chart_series_set_labels(lxw_chart_series *series);
void LXW_CALL chart_series_set_labels_options(lxw_chart_series *series, uint8_t show_name, uint8_t show_category, uint8_t show_value);
lxw_error LXW_CALL chart_series_set_labels_custom(lxw_chart_series *series, lxw_chart_data_label *data_labels[]);
void LXW_CALL chart_series_set_labels_separator(lxw_chart_series *series, uint8_t separator);
void LXW_CALL chart_series_set_labels_position(lxw_chart_series *series, uint8_t position);
void LXW_CALL chart_series_set_labels_leader_line(lxw_chart_series *series);
void LXW_CALL chart_series_set_labels_legend(lxw_chart_series *series);
void LXW_CALL chart_series_set_labels_percentage(lxw_chart_series *series);
void LXW_CALL chart_series_set_labels_num_format(lxw_chart_series *series, const char *num_format);
void LXW_CALL chart_series_set_labels_font(lxw_chart_series *series, lxw_chart_font *font);
void LXW_CALL chart_series_set_labels_line(lxw_chart_series *series, lxw_chart_line *line);
void LXW_CALL chart_series_set_labels_fill(lxw_chart_series *series, lxw_chart_fill *fill);
void LXW_CALL chart_series_set_labels_pattern(lxw_chart_series *series, lxw_chart_pattern *pattern);
void LXW_CALL chart_series_set_trendline(lxw_chart_series *series, uint8_t type, uint8_t value);
void LXW_CALL chart_series_set_trendline_forecast(lxw_chart_series *series, double forward, double backward);
void LXW_CALL chart_series_set_trendline_equation(lxw_chart_series *series);
void LXW_CALL chart_series_set_trendline_r_squared(lxw_chart_series *series);
void LXW_CALL chart_series_set_trendline_intercept(lxw_chart_series *series, double intercept);
void LXW_CALL chart_series_set_trendline_name(lxw_chart_series *series, const char *name);
void LXW_CALL chart_series_set_trendline_line(lxw_chart_series *series, lxw_chart_line *line);
void LXW_CALL chart_series_set_error_bars(lxw_series_error_bars *error_bars, uint8_t type, double value);
void LXW_CALL chart_series_set_error_bars_direction(lxw_series_error_bars *error_bars, uint8_t direction);
void LXW_CALL chart_series_set_error_bars_endcap(lxw_series_error_bars *error_bars, uint8_t endcap);
void LXW_CALL chart_series_set_error_bars_line(lxw_series_error_bars *error_bars, lxw_chart_line *line);

void LXW_CALL chart_axis_set_name(lxw_chart_axis *axis, const char *name);
void LXW_CALL chart_axis_set_name_range(lxw_chart_axis *axis, const char *sheetname, lxw_row_t row, lxw_col_t col);
void LXW_CALL chart_axis_set_name_layout(lxw_chart_axis *axis, lxw_chart_layout *layout);
void LXW_CALL chart_axis_set_name_font(lxw_chart_axis *axis, lxw_chart_font *font);
void LXW_CALL chart_axis_set_num_font(lxw_chart_axis *axis, lxw_chart_font *font);
void LXW_CALL chart_axis_set_num_format(lxw_chart_axis *axis, const char *num_format);
void LXW_CALL chart_axis_set_line(lxw_chart_axis *axis, lxw_chart_line *line);
void LXW_CALL chart_axis_set_fill(lxw_chart_axis *axis, lxw_chart_fill *fill);
void LXW_CALL chart_axis_set_pattern(lxw_chart_axis *axis, lxw_chart_pattern *pattern);
void LXW_CALL chart_axis_set_reverse(lxw_chart_axis *axis);
void LXW_CALL chart_axis_set_crossing(lxw_chart_axis *axis, double value);
void LXW_CALL chart_axis_set_crossing_max(lxw_chart_axis *axis);
void LXW_CALL chart_axis_set_crossing_min(lxw_chart_axis *axis);
void LXW_CALL chart_axis_off(lxw_chart_axis *axis);
void LXW_CALL chart_axis_set_position(lxw_chart_axis *axis, uint8_t position);
void LXW_CALL chart_axis_set_label_position(lxw_chart_axis *axis, uint8_t position);
void LXW_CALL chart_axis_set_label_align(lxw_chart_axis *axis, uint8_t align);
void LXW_CALL chart_axis_set_min(lxw_chart_axis *axis, double min);
void LXW_CALL chart_axis_set_max(lxw_chart_axis *axis, double max);
void LXW_CALL chart_axis_set_log_base(lxw_chart_axis *axis, uint16_t log_base);
void LXW_CALL chart_axis_set_major_tick_mark(lxw_chart_axis *axis, uint8_t type);
void LXW_CALL chart_axis_set_minor_tick_mark(lxw_chart_axis *axis, uint8_t type);
void LXW_CALL chart_axis_set_interval_unit(lxw_chart_axis *axis, uint16_t unit);
void LXW_CALL chart_axis_set_interval_tick(lxw_chart_axis *axis, uint16_t unit);
void LXW_CALL chart_axis_set_major_unit(lxw_chart_axis *axis, double unit);
void LXW_CALL chart_axis_set_minor_unit(lxw_chart_axis *axis, double unit);
void LXW_CALL chart_axis_set_display_units(lxw_chart_axis *axis, uint8_t units);
void LXW_CALL chart_axis_set_display_units_visible(lxw_chart_axis *axis, uint8_t visible);
void LXW_CALL chart_axis_major_gridlines_set_visible(lxw_chart_axis *axis, uint8_t visible);
void LXW_CALL chart_axis_minor_gridlines_set_visible(lxw_chart_axis *axis, uint8_t visible);
void LXW_CALL chart_axis_major_gridlines_set_line(lxw_chart_axis *axis, lxw_chart_line *line);
void LXW_CALL chart_axis_minor_gridlines_set_line(lxw_chart_axis *axis, lxw_chart_line *line);

void LXW_CALL chart_title_set_name(lxw_chart *chart, const char *name);
void LXW_CALL chart_title_set_name_range(lxw_chart *chart, const char *sheetname, lxw_row_t row, lxw_col_t col);
void LXW_CALL chart_title_set_name_font(lxw_chart *chart, lxw_chart_font *font);
void LXW_CALL chart_title_off(lxw_chart *chart);
void LXW_CALL chart_title_set_layout(lxw_chart *chart, lxw_chart_layout *layout);
void LXW_CALL chart_title_set_overlay(lxw_chart *chart, uint8_t overlay);
void LXW_CALL chart_legend_set_position(lxw_chart *chart, uint8_t position);
void LXW_CALL chart_legend_set_layout(lxw_chart *chart, lxw_chart_layout *layout);
void LXW_CALL chart_legend_set_font(lxw_chart *chart, lxw_chart_font *font);
lxw_error LXW_CALL chart_legend_delete_series(lxw_chart *chart, int16_t delete_series[]);
void LXW_CALL chart_chartarea_set_line(lxw_chart *chart, lxw_chart_line *line);
void LXW_CALL chart_chartarea_set_fill(lxw_chart *chart, lxw_chart_fill *fill);
void LXW_CALL chart_chartarea_set_pattern(lxw_chart *chart, lxw_chart_pattern *pattern);
void LXW_CALL chart_plotarea_set_line(lxw_chart *chart, lxw_chart_line *line);
void LXW_CALL chart_plotarea_set_fill(lxw_chart *chart, lxw_chart_fill *fill);
void LXW_CALL chart_plotarea_set_pattern(lxw_chart *chart, lxw_chart_pattern *pattern);
void LXW_CALL chart_plotarea_set_layout(lxw_chart *chart, lxw_chart_layout *layout);
void LXW_CALL chart_set_style(lxw_chart *chart, uint8_t style_id);
void LXW_CALL chart_set_table(lxw_chart *chart);
void LXW_CALL chart_set_table_grid(lxw_chart *chart, uint8_t horizontal, uint8_t vertical, uint8_t outline, uint8_t legend_keys);
void LXW_CALL chart_set_table_font(lxw_chart *chart, lxw_chart_font *font);
void LXW_CALL chart_set_up_down_bars(lxw_chart *chart);
void LXW_CALL chart_set_up_down_bars_format(lxw_chart *chart, lxw_chart_line *up_bar_line, lxw_chart_fill *up_bar_fill, lxw_chart_line *down_bar_line, lxw_chart_fill *down_bar_fill);
void LXW_CALL chart_set_drop_lines(lxw_chart *chart, lxw_chart_line *line);
void LXW_CALL chart_set_high_low_lines(lxw_chart *chart, lxw_chart_line *line);
void LXW_CALL chart_set_series_overlap(lxw_chart *chart, int8_t overlap);
void LXW_CALL chart_set_series_gap(lxw_chart *chart, uint16_t gap);
void LXW_CALL chart_show_blanks_as(lxw_chart *chart, uint8_t option);
void LXW_CALL chart_show_hidden_data(lxw_chart *chart);
void LXW_CALL chart_set_rotation(lxw_chart *chart, uint16_t rotation);
void LXW_CALL chart_set_hole_size(lxw_chart *chart, uint8_t size);

/* ============================================================================
 * FORMAT FUNCTIONS
 * ============================================================================ */

void LXW_CALL format_set_font_name(lxw_format *format, const char *font_name);
void LXW_CALL format_set_font_size(lxw_format *format, double size);
void LXW_CALL format_set_font_color(lxw_format *format, lxw_color_t color);
void LXW_CALL format_set_bold(lxw_format *format);
void LXW_CALL format_set_italic(lxw_format *format);
void LXW_CALL format_set_underline(lxw_format *format, uint8_t style);
void LXW_CALL format_set_font_strikeout(lxw_format *format);
void LXW_CALL format_set_font_script(lxw_format *format, uint8_t style);
void LXW_CALL format_set_font_family(lxw_format *format, uint8_t value);
void LXW_CALL format_set_font_charset(lxw_format *format, uint8_t value);
void LXW_CALL format_set_num_format(lxw_format *format, const char *num_format);
void LXW_CALL format_set_num_format_index(lxw_format *format, uint8_t index);
void LXW_CALL format_set_unlocked(lxw_format *format);
void LXW_CALL format_set_hidden(lxw_format *format);
void LXW_CALL format_set_align(lxw_format *format, uint8_t alignment);
void LXW_CALL format_set_text_wrap(lxw_format *format);
void LXW_CALL format_set_rotation(lxw_format *format, int16_t angle);
void LXW_CALL format_set_indent(lxw_format *format, uint8_t level);
void LXW_CALL format_set_shrink(lxw_format *format);
void LXW_CALL format_set_pattern(lxw_format *format, uint8_t index);
void LXW_CALL format_set_bg_color(lxw_format *format, lxw_color_t color);
void LXW_CALL format_set_fg_color(lxw_format *format, lxw_color_t color);
void LXW_CALL format_set_border(lxw_format *format, uint8_t style);
void LXW_CALL format_set_bottom(lxw_format *format, uint8_t style);
void LXW_CALL format_set_top(lxw_format *format, uint8_t style);
void LXW_CALL format_set_left(lxw_format *format, uint8_t style);
void LXW_CALL format_set_right(lxw_format *format, uint8_t style);
void LXW_CALL format_set_border_color(lxw_format *format, lxw_color_t color);
void LXW_CALL format_set_bottom_color(lxw_format *format, lxw_color_t color);
void LXW_CALL format_set_top_color(lxw_format *format, lxw_color_t color);
void LXW_CALL format_set_left_color(lxw_format *format, lxw_color_t color);
void LXW_CALL format_set_right_color(lxw_format *format, lxw_color_t color);
void LXW_CALL format_set_diag_type(lxw_format *format, uint8_t type);
void LXW_CALL format_set_diag_border(lxw_format *format, uint8_t style);
void LXW_CALL format_set_diag_color(lxw_format *format, lxw_color_t color);
void LXW_CALL format_set_quote_prefix(lxw_format *format);

/* ============================================================================
 * UTILITY FUNCTIONS
 * ============================================================================ */

const char * LXW_CALL lxw_version(void);
uint16_t LXW_CALL lxw_version_id(void);
char * LXW_CALL lxw_strerror(lxw_error error_num);

void LXW_CALL lxw_col_to_name(char *col_name, lxw_col_t col_num, uint8_t absolute);
void LXW_CALL lxw_rowcol_to_cell(char *cell_name, lxw_row_t row, lxw_col_t col);
void LXW_CALL lxw_rowcol_to_cell_abs(char *cell_name, lxw_row_t row, lxw_col_t col, uint8_t abs_row, uint8_t abs_col);
void LXW_CALL lxw_rowcol_to_range(char *range, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void LXW_CALL lxw_rowcol_to_range_abs(char *range, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void LXW_CALL lxw_rowcol_to_formula_abs(char *formula, const char *sheetname, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);

uint32_t LXW_CALL lxw_name_to_row(const char *row_str);
uint16_t LXW_CALL lxw_name_to_col(const char *col_str);
uint32_t LXW_CALL lxw_name_to_row_2(const char *row_str);
uint16_t LXW_CALL lxw_name_to_col_2(const char *col_str);

double LXW_CALL lxw_datetime_to_excel_datetime(lxw_datetime *datetime);
double LXW_CALL lxw_datetime_to_excel_date_with_epoch(lxw_datetime *datetime, uint8_t use_1904_epoch);
double LXW_CALL lxw_unixtime_to_excel_date(int64_t unixtime);
double LXW_CALL lxw_unixtime_to_excel_date_with_epoch(int64_t unixtime, uint8_t use_1904_epoch);

/* ============================================================================
 * MACRO REPLACEMENT FUNCTIONS
 *
 * The following inline functions replace the original macros:
 * - CELL(cell) -> use lxw_name_to_row(cell) and lxw_name_to_col(cell) separately
 * - COLS(cols) -> use lxw_name_to_col(cols) and lxw_name_to_col_2(cols) separately
 * - RANGE(range) -> use lxw_name_to_row(range), lxw_name_to_col(range),
 *                   lxw_name_to_row_2(range), lxw_name_to_col_2(range) separately
 *
 * Example usage:
 *   Instead of: worksheet_write_string(worksheet, CELL("A1"), "Hello", NULL);
 *   Use:        worksheet_write_string(worksheet, lxw_name_to_row("A1"),
 *                                      lxw_name_to_col("A1"), "Hello", NULL);
 *
 *   Instead of: worksheet_set_column(worksheet, COLS("B:D"), 20, NULL, NULL);
 *   Use:        worksheet_set_column(worksheet, lxw_name_to_col("B:D"),
 *                                    lxw_name_to_col_2("B:D"), 20, NULL, NULL);
 *
 *   Instead of: worksheet_print_area(worksheet, RANGE("A1:K42"));
 *   Use:        worksheet_print_area(worksheet, lxw_name_to_row("A1:K42"),
 *                                    lxw_name_to_col("A1:K42"),
 *                                    lxw_name_to_row_2("A1:K42"),
 *                                    lxw_name_to_col_2("A1:K42"));
 * ============================================================================ */

#ifdef __cplusplus
}
#endif

#endif /* __LIBXLSXWRITER_LV_H__ */
