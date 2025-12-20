/*
 * libxlsxwriter_LV.h - LabVIEW-compatible header for libxlsxwriter
 *
 * This header is specifically designed for LabVIEW's Import Shared Library wizard.
 * All preprocessor conditionals have been removed for compatibility.
 * All pointer types use unsigned long (32-bit handle).
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 */

#ifndef __LIBXLSXWRITER_LV_H__
#define __LIBXLSXWRITER_LV_H__

/* ============================================================================
 * Basic Type Definitions (no preprocessor conditionals)
 * ============================================================================ */

typedef signed char        int8_t;
typedef signed short       int16_t;
typedef signed int         int32_t;
typedef unsigned char      uint8_t;
typedef unsigned short     uint16_t;
typedef unsigned int       uint32_t;
typedef unsigned int       size_t;

/* ============================================================================
 * Library-Specific Type Definitions
 * ============================================================================ */

typedef uint32_t lxw_row_t;
typedef uint16_t lxw_col_t;
typedef uint32_t lxw_color_t;

/* ============================================================================
 * Opaque Handle Types (use unsigned long for 32-bit pointers)
 *
 * In LabVIEW, map these to "Unsigned 32-bit Integer" (U32)
 * ============================================================================ */

typedef unsigned long lxw_workbook;
typedef unsigned long lxw_worksheet;
typedef unsigned long lxw_chartsheet;
typedef unsigned long lxw_chart;
typedef unsigned long lxw_chart_series;
typedef unsigned long lxw_chart_axis;
typedef unsigned long lxw_format;
typedef unsigned long lxw_series_error_bars;
typedef unsigned long lxw_styles;
typedef unsigned long lxw_relationships;
typedef unsigned long lxw_drawing;
typedef unsigned long lxw_file_handle;

/* ============================================================================
 * Error Codes
 * ============================================================================ */

typedef enum lxw_error {
    LXW_NO_ERROR = 0,
    LXW_ERROR_MEMORY_MALLOC_FAILED = 1,
    LXW_ERROR_CREATING_XLSX_FILE = 2,
    LXW_ERROR_CREATING_TMPFILE = 3,
    LXW_ERROR_READING_TMPFILE = 4,
    LXW_ERROR_ZIP_FILE_OPERATION = 5,
    LXW_ERROR_ZIP_PARAMETER_ERROR = 6,
    LXW_ERROR_ZIP_BAD_ZIP_FILE = 7,
    LXW_ERROR_ZIP_INTERNAL_ERROR = 8,
    LXW_ERROR_ZIP_FILE_ADD = 9,
    LXW_ERROR_ZIP_CLOSE = 10,
    LXW_ERROR_FEATURE_NOT_SUPPORTED = 11,
    LXW_ERROR_NULL_PARAMETER_IGNORED = 12,
    LXW_ERROR_PARAMETER_VALIDATION = 13,
    LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE = 20
} lxw_error;

/* ============================================================================
 * Simple Structures (LabVIEW can create clusters for these)
 * ============================================================================ */

typedef struct lxw_datetime {
    int32_t year;
    int32_t month;
    int32_t day;
    int32_t hour;
    int32_t min;
    double sec;
} lxw_datetime;

typedef struct lxw_chart_line {
    lxw_color_t color;
    uint8_t none;
    float width;
    uint8_t dash_type;
    uint8_t transparency;
} lxw_chart_line;

typedef struct lxw_chart_fill {
    lxw_color_t color;
    uint8_t none;
    uint8_t transparency;
} lxw_chart_fill;

typedef struct lxw_chart_pattern {
    lxw_color_t fg_color;
    lxw_color_t bg_color;
    uint8_t type;
} lxw_chart_pattern;

typedef struct lxw_chart_font {
    unsigned long name;
    double size;
    uint8_t bold;
    uint8_t italic;
    uint8_t underline;
    int32_t rotation;
    lxw_color_t color;
    uint8_t pitch_family;
    uint8_t charset;
    int8_t baseline;
} lxw_chart_font;

typedef struct lxw_chart_point {
    unsigned long line;
    unsigned long fill;
    unsigned long pattern;
} lxw_chart_point;

typedef struct lxw_image_options {
    int32_t x_offset;
    int32_t y_offset;
    double x_scale;
    double y_scale;
    uint32_t row;
    uint16_t col;
    unsigned long url;
    unsigned long tip;
    uint8_t object_position;
    unsigned long description;
    unsigned long decorative;
} lxw_image_options;

typedef struct lxw_chart_options {
    int32_t x_offset;
    int32_t y_offset;
    double x_scale;
    double y_scale;
    uint8_t object_position;
    unsigned long description;
    unsigned long decorative;
} lxw_chart_options;

typedef struct lxw_row_col_options {
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
} lxw_row_col_options;

typedef struct lxw_protection {
    uint8_t no_select_locked_cells;
    uint8_t no_select_unlocked_cells;
    uint8_t format_cells;
    uint8_t format_columns;
    uint8_t format_rows;
    uint8_t insert_columns;
    uint8_t insert_rows;
    uint8_t insert_hyperlinks;
    uint8_t delete_columns;
    uint8_t delete_rows;
    uint8_t sort;
    uint8_t autofilter;
    uint8_t pivot_tables;
    uint8_t scenarios;
    uint8_t objects;
    uint8_t no_content;
    uint8_t no_objects;
} lxw_protection;

typedef struct lxw_header_footer_options {
    double margin;
    unsigned long image_left;
    unsigned long image_center;
    unsigned long image_right;
} lxw_header_footer_options;

/* ============================================================================
 * Enumeration Constants (as defines for LabVIEW compatibility)
 * ============================================================================ */

#define LXW_CHART_NONE                          0
#define LXW_CHART_AREA                          1
#define LXW_CHART_AREA_STACKED                  2
#define LXW_CHART_AREA_STACKED_PERCENT          3
#define LXW_CHART_BAR                           4
#define LXW_CHART_BAR_STACKED                   5
#define LXW_CHART_BAR_STACKED_PERCENT           6
#define LXW_CHART_COLUMN                        7
#define LXW_CHART_COLUMN_STACKED                8
#define LXW_CHART_COLUMN_STACKED_PERCENT        9
#define LXW_CHART_DOUGHNUT                      10
#define LXW_CHART_LINE                          11
#define LXW_CHART_LINE_STACKED                  12
#define LXW_CHART_LINE_STACKED_PERCENT          13
#define LXW_CHART_PIE                           14
#define LXW_CHART_SCATTER                       15
#define LXW_CHART_SCATTER_STRAIGHT              16
#define LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS 17
#define LXW_CHART_SCATTER_SMOOTH                18
#define LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS   19
#define LXW_CHART_RADAR                         20
#define LXW_CHART_RADAR_WITH_MARKERS            21
#define LXW_CHART_RADAR_FILLED                  22

#define LXW_CHART_LEGEND_NONE                   0
#define LXW_CHART_LEGEND_RIGHT                  1
#define LXW_CHART_LEGEND_LEFT                   2
#define LXW_CHART_LEGEND_TOP                    3
#define LXW_CHART_LEGEND_BOTTOM                 4
#define LXW_CHART_LEGEND_TOP_RIGHT              5
#define LXW_CHART_LEGEND_OVERLAY_RIGHT          6
#define LXW_CHART_LEGEND_OVERLAY_LEFT           7

#define LXW_CHART_MARKER_AUTOMATIC              0
#define LXW_CHART_MARKER_NONE                   1
#define LXW_CHART_MARKER_SQUARE                 2
#define LXW_CHART_MARKER_DIAMOND                3
#define LXW_CHART_MARKER_TRIANGLE               4
#define LXW_CHART_MARKER_X                      5
#define LXW_CHART_MARKER_STAR                   6
#define LXW_CHART_MARKER_SHORT_DASH             7
#define LXW_CHART_MARKER_LONG_DASH              8
#define LXW_CHART_MARKER_CIRCLE                 9
#define LXW_CHART_MARKER_PLUS                   10

#define LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO   0
#define LXW_CHART_AXIS_LABEL_POSITION_HIGH      1
#define LXW_CHART_AXIS_LABEL_POSITION_LOW       2
#define LXW_CHART_AXIS_LABEL_POSITION_NONE      3

#define LXW_ALIGN_NONE                          0
#define LXW_ALIGN_LEFT                          1
#define LXW_ALIGN_CENTER                        2
#define LXW_ALIGN_RIGHT                         3
#define LXW_ALIGN_FILL                          4
#define LXW_ALIGN_JUSTIFY                       5
#define LXW_ALIGN_CENTER_ACROSS                 6
#define LXW_ALIGN_DISTRIBUTED                   7
#define LXW_ALIGN_VERTICAL_TOP                  8
#define LXW_ALIGN_VERTICAL_BOTTOM               9
#define LXW_ALIGN_VERTICAL_CENTER               10
#define LXW_ALIGN_VERTICAL_JUSTIFY              11
#define LXW_ALIGN_VERTICAL_DISTRIBUTED          12

#define LXW_BORDER_NONE                         0
#define LXW_BORDER_THIN                         1
#define LXW_BORDER_MEDIUM                       2
#define LXW_BORDER_DASHED                       3
#define LXW_BORDER_DOTTED                       4
#define LXW_BORDER_THICK                        5
#define LXW_BORDER_DOUBLE                       6
#define LXW_BORDER_HAIR                         7

#define LXW_PATTERN_NONE                        0
#define LXW_PATTERN_SOLID                       1
#define LXW_PATTERN_MEDIUM_GRAY                 2
#define LXW_PATTERN_DARK_GRAY                   3
#define LXW_PATTERN_LIGHT_GRAY                  4

#define LXW_COLOR_BLACK     0x000000
#define LXW_COLOR_BLUE      0x0000FF
#define LXW_COLOR_BROWN     0x800000
#define LXW_COLOR_CYAN      0x00FFFF
#define LXW_COLOR_GRAY      0x808080
#define LXW_COLOR_GREEN     0x008000
#define LXW_COLOR_LIME      0x00FF00
#define LXW_COLOR_MAGENTA   0xFF00FF
#define LXW_COLOR_NAVY      0x000080
#define LXW_COLOR_ORANGE    0xFF6600
#define LXW_COLOR_PINK      0xFF00FF
#define LXW_COLOR_PURPLE    0x800080
#define LXW_COLOR_RED       0xFF0000
#define LXW_COLOR_SILVER    0xC0C0C0
#define LXW_COLOR_WHITE     0xFFFFFF
#define LXW_COLOR_YELLOW    0xFFFF00

/* ============================================================================
 * Workbook Functions
 * ============================================================================ */

lxw_workbook workbook_new(const char *filename);
lxw_workbook workbook_new_opt(const char *filename, unsigned long options);
lxw_worksheet workbook_add_worksheet(lxw_workbook workbook, const char *sheetname);
lxw_chartsheet workbook_add_chartsheet(lxw_workbook workbook, const char *sheetname);
lxw_format workbook_add_format(lxw_workbook workbook);
lxw_chart workbook_add_chart(lxw_workbook workbook, uint8_t chart_type);
lxw_error workbook_close(lxw_workbook workbook);
lxw_error workbook_set_properties(lxw_workbook workbook, unsigned long properties);
lxw_error workbook_define_name(lxw_workbook workbook, const char *name, const char *formula);
lxw_worksheet workbook_get_worksheet_by_name(lxw_workbook workbook, const char *name);
lxw_chartsheet workbook_get_chartsheet_by_name(lxw_workbook workbook, const char *name);
lxw_error workbook_validate_sheet_name(lxw_workbook workbook, const char *sheetname);
lxw_error workbook_set_custom_property_string(lxw_workbook workbook, const char *name, const char *value);
lxw_error workbook_set_custom_property_number(lxw_workbook workbook, const char *name, double value);
lxw_error workbook_set_custom_property_integer(lxw_workbook workbook, const char *name, int32_t value);
lxw_error workbook_set_custom_property_boolean(lxw_workbook workbook, const char *name, uint8_t value);
lxw_error workbook_set_custom_property_datetime(lxw_workbook workbook, const char *name, lxw_datetime *datetime);

/* ============================================================================
 * Worksheet Functions
 * ============================================================================ */

lxw_error worksheet_write_number(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, double number, lxw_format format);
lxw_error worksheet_write_string(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *string, lxw_format format);
lxw_error worksheet_write_formula(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format);
lxw_error worksheet_write_formula_num(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format, double result);
lxw_error worksheet_write_formula_str(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format, const char *result);
lxw_error worksheet_write_array_formula(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format format);
lxw_error worksheet_write_array_formula_num(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format format, double result);
lxw_error worksheet_write_dynamic_array_formula(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format format);
lxw_error worksheet_write_dynamic_array_formula_num(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format format, double result);
lxw_error worksheet_write_dynamic_formula(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format);
lxw_error worksheet_write_dynamic_formula_num(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format, double result);
lxw_error worksheet_write_datetime(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, lxw_datetime *datetime, lxw_format format);
lxw_error worksheet_write_unixtime(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, int64_t unixtime, lxw_format format);
lxw_error worksheet_write_url(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *url, lxw_format format);
lxw_error worksheet_write_url_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *url, lxw_format format, const char *string, const char *tooltip);
lxw_error worksheet_write_boolean(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, int value, lxw_format format);
lxw_error worksheet_write_blank(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, lxw_format format);
lxw_error worksheet_write_rich_string(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, unsigned long rich_string, lxw_format format);
lxw_error worksheet_write_comment(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *string);
lxw_error worksheet_write_comment_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *string, unsigned long options);

lxw_error worksheet_set_row(lxw_worksheet worksheet, lxw_row_t row, double height, lxw_format format);
lxw_error worksheet_set_row_opt(lxw_worksheet worksheet, lxw_row_t row, double height, lxw_format format, lxw_row_col_options *options);
lxw_error worksheet_set_row_pixels(lxw_worksheet worksheet, lxw_row_t row, uint32_t pixels, lxw_format format);
lxw_error worksheet_set_row_pixels_opt(lxw_worksheet worksheet, lxw_row_t row, uint32_t pixels, lxw_format format, lxw_row_col_options *options);
lxw_error worksheet_set_column(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col, double width, lxw_format format);
lxw_error worksheet_set_column_opt(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col, double width, lxw_format format, lxw_row_col_options *options);
lxw_error worksheet_set_column_pixels(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col, uint32_t pixels, lxw_format format);
lxw_error worksheet_set_column_pixels_opt(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col, uint32_t pixels, lxw_format format, lxw_row_col_options *options);

lxw_error worksheet_insert_image(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename);
lxw_error worksheet_insert_image_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename, lxw_image_options *options);
lxw_error worksheet_insert_image_buffer(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size);
lxw_error worksheet_insert_image_buffer_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size, lxw_image_options *options);
lxw_error worksheet_insert_chart(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, lxw_chart chart);
lxw_error worksheet_insert_chart_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, lxw_chart chart, lxw_chart_options *options);

lxw_error worksheet_merge_range(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *string, lxw_format format);
lxw_error worksheet_autofilter(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
lxw_error worksheet_filter_column(lxw_worksheet worksheet, lxw_col_t col, unsigned long rule);
lxw_error worksheet_filter_column2(lxw_worksheet worksheet, lxw_col_t col, unsigned long rule1, unsigned long rule2, uint8_t and_or);
lxw_error worksheet_filter_list(lxw_worksheet worksheet, lxw_col_t col, unsigned long list);

lxw_error worksheet_data_validation_cell(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, unsigned long validation);
lxw_error worksheet_data_validation_range(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, unsigned long validation);
lxw_error worksheet_conditional_format_cell(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, unsigned long conditional_format);
lxw_error worksheet_conditional_format_range(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, unsigned long conditional_format);

void worksheet_activate(lxw_worksheet worksheet);
void worksheet_select(lxw_worksheet worksheet);
void worksheet_hide(lxw_worksheet worksheet);
void worksheet_set_first_sheet(lxw_worksheet worksheet);
lxw_error worksheet_freeze_panes(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col);
lxw_error worksheet_split_panes(lxw_worksheet worksheet, double vertical, double horizontal);
lxw_error worksheet_freeze_panes_opt(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t top_row, lxw_col_t left_col, uint8_t type);

void worksheet_set_selection(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void worksheet_set_landscape(lxw_worksheet worksheet);
void worksheet_set_portrait(lxw_worksheet worksheet);
void worksheet_set_page_view(lxw_worksheet worksheet);
void worksheet_set_paper(lxw_worksheet worksheet, uint8_t paper_type);
void worksheet_set_margins(lxw_worksheet worksheet, double left, double right, double top, double bottom);
lxw_error worksheet_set_header(lxw_worksheet worksheet, const char *header);
lxw_error worksheet_set_footer(lxw_worksheet worksheet, const char *footer);
lxw_error worksheet_set_header_opt(lxw_worksheet worksheet, const char *header, lxw_header_footer_options *options);
lxw_error worksheet_set_footer_opt(lxw_worksheet worksheet, const char *footer, lxw_header_footer_options *options);
void worksheet_set_h_pagebreaks(lxw_worksheet worksheet, unsigned long breaks);
void worksheet_set_v_pagebreaks(lxw_worksheet worksheet, unsigned long breaks);
void worksheet_print_across(lxw_worksheet worksheet);
void worksheet_set_zoom(lxw_worksheet worksheet, uint16_t scale);
void worksheet_gridlines(lxw_worksheet worksheet, uint8_t option);
void worksheet_center_horizontally(lxw_worksheet worksheet);
void worksheet_center_vertically(lxw_worksheet worksheet);
void worksheet_print_row_col_headers(lxw_worksheet worksheet);
lxw_error worksheet_repeat_rows(lxw_worksheet worksheet, lxw_row_t first_row, lxw_row_t last_row);
lxw_error worksheet_repeat_columns(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col);
lxw_error worksheet_print_area(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void worksheet_fit_to_pages(lxw_worksheet worksheet, uint16_t width, uint16_t height);
void worksheet_set_start_page(lxw_worksheet worksheet, uint16_t start_page);
void worksheet_set_print_scale(lxw_worksheet worksheet, uint16_t scale);
void worksheet_right_to_left(lxw_worksheet worksheet);
void worksheet_hide_zero(lxw_worksheet worksheet);
void worksheet_set_tab_color(lxw_worksheet worksheet, lxw_color_t color);
void worksheet_protect(lxw_worksheet worksheet, const char *password, lxw_protection *options);
void worksheet_outline_settings(lxw_worksheet worksheet, uint8_t visible, uint8_t symbols_below, uint8_t symbols_right, uint8_t auto_style);
void worksheet_set_default_row(lxw_worksheet worksheet, double height, uint8_t hide_unused_rows);
lxw_error worksheet_set_vba_name(lxw_worksheet worksheet, const char *name);
void worksheet_show_comments(lxw_worksheet worksheet);
void worksheet_set_comments_author(lxw_worksheet worksheet, const char *author);
void worksheet_ignore_errors(lxw_worksheet worksheet, uint8_t type, const char *range);
lxw_error worksheet_set_background(lxw_worksheet worksheet, const char *filename);
lxw_error worksheet_set_background_buffer(lxw_worksheet worksheet, const unsigned char *image_buffer, size_t image_size);

/* ============================================================================
 * Chartsheet Functions
 * ============================================================================ */

lxw_error chartsheet_set_chart(lxw_chartsheet chartsheet, lxw_chart chart);
lxw_error chartsheet_set_chart_opt(lxw_chartsheet chartsheet, lxw_chart chart, lxw_chart_options *options);
void chartsheet_activate(lxw_chartsheet chartsheet);
void chartsheet_select(lxw_chartsheet chartsheet);
void chartsheet_hide(lxw_chartsheet chartsheet);
void chartsheet_set_first_sheet(lxw_chartsheet chartsheet);
void chartsheet_set_tab_color(lxw_chartsheet chartsheet, lxw_color_t color);
void chartsheet_protect(lxw_chartsheet chartsheet, const char *password, lxw_protection *options);
void chartsheet_set_zoom(lxw_chartsheet chartsheet, uint16_t scale);
void chartsheet_set_landscape(lxw_chartsheet chartsheet);
void chartsheet_set_portrait(lxw_chartsheet chartsheet);
void chartsheet_set_paper(lxw_chartsheet chartsheet, uint8_t paper_type);
void chartsheet_set_margins(lxw_chartsheet chartsheet, double left, double right, double top, double bottom);
lxw_error chartsheet_set_header(lxw_chartsheet chartsheet, const char *header);
lxw_error chartsheet_set_footer(lxw_chartsheet chartsheet, const char *footer);
lxw_error chartsheet_set_header_opt(lxw_chartsheet chartsheet, const char *header, lxw_header_footer_options *options);
lxw_error chartsheet_set_footer_opt(lxw_chartsheet chartsheet, const char *footer, lxw_header_footer_options *options);

/* ============================================================================
 * Chart Functions
 * ============================================================================ */

lxw_chart_series chart_add_series(lxw_chart chart, const char *categories, const char *values);
lxw_chart_series chart_add_series_on_axis(lxw_chart chart, const char *categories, const char *values, lxw_chart_axis y_axis);
void chart_series_set_categories(lxw_chart_series series, const char *sheetname, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void chart_series_set_values(lxw_chart_series series, const char *sheetname, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void chart_series_set_name(lxw_chart_series series, const char *name);
void chart_series_set_name_range(lxw_chart_series series, const char *sheetname, lxw_row_t row, lxw_col_t col);
void chart_series_set_line(lxw_chart_series series, lxw_chart_line *line);
void chart_series_set_fill(lxw_chart_series series, lxw_chart_fill *fill);
void chart_series_set_invert_if_negative(lxw_chart_series series);
void chart_series_set_pattern(lxw_chart_series series, lxw_chart_pattern *pattern);
void chart_series_set_marker_type(lxw_chart_series series, uint8_t type);
void chart_series_set_marker_size(lxw_chart_series series, uint8_t size);
void chart_series_set_marker_line(lxw_chart_series series, lxw_chart_line *line);
void chart_series_set_marker_fill(lxw_chart_series series, lxw_chart_fill *fill);
void chart_series_set_marker_pattern(lxw_chart_series series, lxw_chart_pattern *pattern);
void chart_series_set_points(lxw_chart_series series, unsigned long points);
void chart_series_set_smooth(lxw_chart_series series, uint8_t smooth);
void chart_series_set_labels(lxw_chart_series series);
void chart_series_set_labels_options(lxw_chart_series series, uint8_t show_name, uint8_t show_category, uint8_t show_value);
void chart_series_set_labels_separator(lxw_chart_series series, uint8_t separator);
void chart_series_set_labels_position(lxw_chart_series series, uint8_t position);
void chart_series_set_labels_leader_line(lxw_chart_series series);
void chart_series_set_labels_legend(lxw_chart_series series);
void chart_series_set_labels_percentage(lxw_chart_series series);
void chart_series_set_labels_num_format(lxw_chart_series series, const char *num_format);
void chart_series_set_labels_font(lxw_chart_series series, lxw_chart_font *font);
void chart_series_set_labels_line(lxw_chart_series series, lxw_chart_line *line);
void chart_series_set_labels_fill(lxw_chart_series series, lxw_chart_fill *fill);
void chart_series_set_labels_pattern(lxw_chart_series series, lxw_chart_pattern *pattern);
lxw_error chart_series_set_trendline(lxw_chart_series series, uint8_t type, uint8_t value);
lxw_error chart_series_set_trendline_forecast(lxw_chart_series series, double forward, double backward);
lxw_error chart_series_set_trendline_equation(lxw_chart_series series);
lxw_error chart_series_set_trendline_r_squared(lxw_chart_series series);
lxw_error chart_series_set_trendline_intercept(lxw_chart_series series, double intercept);
lxw_error chart_series_set_trendline_name(lxw_chart_series series, const char *name);
lxw_error chart_series_set_trendline_line(lxw_chart_series series, lxw_chart_line *line);
lxw_series_error_bars chart_series_get_error_bars(lxw_chart_series series, uint8_t axis_type);
void chart_series_set_error_bars(lxw_series_error_bars error_bars, uint8_t type, double value);
void chart_series_set_error_bars_direction(lxw_series_error_bars error_bars, uint8_t direction);
void chart_series_set_error_bars_endcap(lxw_series_error_bars error_bars, uint8_t endcap);
void chart_series_set_error_bars_line(lxw_series_error_bars error_bars, lxw_chart_line *line);

lxw_chart_axis chart_axis_get(lxw_chart chart, uint8_t axis_type);
void chart_axis_set_name(lxw_chart_axis axis, const char *name);
void chart_axis_set_name_range(lxw_chart_axis axis, const char *sheetname, lxw_row_t row, lxw_col_t col);
void chart_axis_set_name_font(lxw_chart_axis axis, lxw_chart_font *font);
void chart_axis_set_num_font(lxw_chart_axis axis, lxw_chart_font *font);
void chart_axis_set_num_format(lxw_chart_axis axis, const char *num_format);
void chart_axis_set_line(lxw_chart_axis axis, lxw_chart_line *line);
void chart_axis_set_fill(lxw_chart_axis axis, lxw_chart_fill *fill);
void chart_axis_set_pattern(lxw_chart_axis axis, lxw_chart_pattern *pattern);
void chart_axis_set_reverse(lxw_chart_axis axis);
void chart_axis_set_crossing(lxw_chart_axis axis, double value);
void chart_axis_set_crossing_max(lxw_chart_axis axis);
void chart_axis_set_crossing_min(lxw_chart_axis axis);
void chart_axis_off(lxw_chart_axis axis);
void chart_axis_set_position(lxw_chart_axis axis, uint8_t position);
void chart_axis_set_label_position(lxw_chart_axis axis, uint8_t position);
void chart_axis_set_label_align(lxw_chart_axis axis, uint8_t align);
void chart_axis_set_min(lxw_chart_axis axis, double min);
void chart_axis_set_max(lxw_chart_axis axis, double max);
void chart_axis_set_log_base(lxw_chart_axis axis, uint16_t log_base);
void chart_axis_set_major_tick_mark(lxw_chart_axis axis, uint8_t type);
void chart_axis_set_minor_tick_mark(lxw_chart_axis axis, uint8_t type);
void chart_axis_set_interval_unit(lxw_chart_axis axis, uint16_t unit);
void chart_axis_set_interval_tick(lxw_chart_axis axis, uint16_t unit);
void chart_axis_set_major_unit(lxw_chart_axis axis, double unit);
void chart_axis_set_minor_unit(lxw_chart_axis axis, double unit);
void chart_axis_set_display_units(lxw_chart_axis axis, uint8_t units);
void chart_axis_set_display_units_visible(lxw_chart_axis axis, uint8_t visible);
void chart_axis_major_gridlines_set_visible(lxw_chart_axis axis, uint8_t visible);
void chart_axis_minor_gridlines_set_visible(lxw_chart_axis axis, uint8_t visible);
void chart_axis_major_gridlines_set_line(lxw_chart_axis axis, lxw_chart_line *line);
void chart_axis_minor_gridlines_set_line(lxw_chart_axis axis, lxw_chart_line *line);

void chart_title_set_name(lxw_chart chart, const char *name);
void chart_title_set_name_range(lxw_chart chart, const char *sheetname, lxw_row_t row, lxw_col_t col);
void chart_title_set_name_font(lxw_chart chart, lxw_chart_font *font);
void chart_title_off(lxw_chart chart);
void chart_legend_set_position(lxw_chart chart, uint8_t position);
void chart_legend_set_font(lxw_chart chart, lxw_chart_font *font);
lxw_error chart_legend_delete_series(lxw_chart chart, int16_t delete_series);
void chart_chartarea_set_line(lxw_chart chart, lxw_chart_line *line);
void chart_chartarea_set_fill(lxw_chart chart, lxw_chart_fill *fill);
void chart_chartarea_set_pattern(lxw_chart chart, lxw_chart_pattern *pattern);
void chart_plotarea_set_line(lxw_chart chart, lxw_chart_line *line);
void chart_plotarea_set_fill(lxw_chart chart, lxw_chart_fill *fill);
void chart_plotarea_set_pattern(lxw_chart chart, lxw_chart_pattern *pattern);
void chart_set_style(lxw_chart chart, uint8_t style_id);
void chart_set_table(lxw_chart chart);
void chart_set_table_grid(lxw_chart chart, uint8_t horizontal, uint8_t vertical, uint8_t outline, uint8_t legend_keys);
void chart_set_table_font(lxw_chart chart, lxw_chart_font *font);
void chart_set_up_down_bars(lxw_chart chart);
void chart_set_up_down_bars_format(lxw_chart chart, lxw_chart_line *up_bar_line, lxw_chart_fill *up_bar_fill, lxw_chart_line *down_bar_line, lxw_chart_fill *down_bar_fill);
void chart_set_drop_lines(lxw_chart chart, lxw_chart_line *line);
void chart_set_high_low_lines(lxw_chart chart, lxw_chart_line *line);
void chart_set_series_overlap(lxw_chart chart, int8_t overlap);
void chart_set_series_gap(lxw_chart chart, uint16_t gap);
void chart_show_blanks_as(lxw_chart chart, uint8_t option);
void chart_show_hidden_data(lxw_chart chart);
void chart_set_rotation(lxw_chart chart, uint16_t rotation);
void chart_set_hole_size(lxw_chart chart, uint8_t size);

/* Axis accessor functions for LabVIEW */
lxw_chart_axis chart_get_x_axis(lxw_chart chart);
lxw_chart_axis chart_get_y_axis(lxw_chart chart);
lxw_chart_axis chart_get_y2_axis(lxw_chart chart);

/* ============================================================================
 * Format Functions
 * ============================================================================ */

void format_set_font_name(lxw_format format, const char *font_name);
void format_set_font_size(lxw_format format, double size);
void format_set_font_color(lxw_format format, lxw_color_t color);
void format_set_bold(lxw_format format);
void format_set_italic(lxw_format format);
void format_set_underline(lxw_format format, uint8_t style);
void format_set_font_strikeout(lxw_format format);
void format_set_font_script(lxw_format format, uint8_t style);
void format_set_num_format(lxw_format format, const char *num_format);
void format_set_num_format_index(lxw_format format, uint8_t index);
void format_set_unlocked(lxw_format format);
void format_set_hidden(lxw_format format);
void format_set_align(lxw_format format, uint8_t alignment);
void format_set_text_wrap(lxw_format format);
void format_set_rotation(lxw_format format, int16_t angle);
void format_set_indent(lxw_format format, uint8_t level);
void format_set_shrink(lxw_format format);
void format_set_pattern(lxw_format format, uint8_t pattern);
void format_set_bg_color(lxw_format format, lxw_color_t color);
void format_set_fg_color(lxw_format format, lxw_color_t color);
void format_set_border(lxw_format format, uint8_t style);
void format_set_bottom(lxw_format format, uint8_t style);
void format_set_top(lxw_format format, uint8_t style);
void format_set_left(lxw_format format, uint8_t style);
void format_set_right(lxw_format format, uint8_t style);
void format_set_border_color(lxw_format format, lxw_color_t color);
void format_set_bottom_color(lxw_format format, lxw_color_t color);
void format_set_top_color(lxw_format format, lxw_color_t color);
void format_set_left_color(lxw_format format, lxw_color_t color);
void format_set_right_color(lxw_format format, lxw_color_t color);
void format_set_diag_type(lxw_format format, uint8_t value);
void format_set_diag_border(lxw_format format, uint8_t value);
void format_set_diag_color(lxw_format format, lxw_color_t color);
void format_set_font_outline(lxw_format format);
void format_set_font_shadow(lxw_format format);
void format_set_font_family(lxw_format format, uint8_t value);
void format_set_font_charset(lxw_format format, uint8_t value);
void format_set_font_scheme(lxw_format format, const char *font_scheme);
void format_set_font_condense(lxw_format format);
void format_set_font_extend(lxw_format format);
void format_set_reading_order(lxw_format format, uint8_t value);
void format_set_theme(lxw_format format, uint8_t value);
void format_set_hyperlink(lxw_format format);
void format_set_color_indexed(lxw_format format, uint8_t value);
void format_set_font_only(lxw_format format);

/* ============================================================================
 * Utility Functions
 * ============================================================================ */

void lxw_version(void);
const char* lxw_version_id(void);
const char* lxw_strerror(lxw_error error_num);
double lxw_datetime_to_excel_datetime(lxw_datetime *datetime);
int32_t lxw_unixtime_to_excel_date(int64_t unixtime);
double lxw_unixtime_to_excel_date_epoch(int64_t unixtime, uint8_t is_date_1904);

#endif /* __LIBXLSXWRITER_LV_H__ */
