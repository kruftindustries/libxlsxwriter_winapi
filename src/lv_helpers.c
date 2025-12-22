/*
 * lv_helpers.c - LabVIEW helper functions for libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 */

/* Required for strdup() on POSIX systems */
#define _POSIX_C_SOURCE 200809L

#include <stdlib.h>
#include <string.h>
#include "xlsxwriter.h"

#ifdef _WIN32
#include <windows.h>
#endif

/*
 * Convert an ANSI string to UTF-8.
 * Returns a newly allocated string that must be freed by the caller.
 * Returns NULL if conversion fails or input is NULL.
 */
static char *
ansi_to_utf8(const char *ansi_str)
{
    if (!ansi_str)
        return NULL;

#ifdef _WIN32
    int wide_len;
    int utf8_len;
    wchar_t *wide_str;
    char *utf8_str;

    /* Get the length needed for the wide string */
    wide_len = MultiByteToWideChar(CP_ACP, 0, ansi_str, -1, NULL, 0);
    if (wide_len == 0)
        return NULL;

    wide_str = (wchar_t *) malloc(wide_len * sizeof(wchar_t));
    if (!wide_str)
        return NULL;

    /* Convert ANSI to wide string */
    if (MultiByteToWideChar(CP_ACP, 0, ansi_str, -1, wide_str, wide_len) == 0) {
        free(wide_str);
        return NULL;
    }

    /* Get the length needed for UTF-8 string */
    utf8_len =
        WideCharToMultiByte(CP_UTF8, 0, wide_str, -1, NULL, 0, NULL, NULL);
    if (utf8_len == 0) {
        free(wide_str);
        return NULL;
    }

    utf8_str = (char *) malloc(utf8_len);
    if (!utf8_str) {
        free(wide_str);
        return NULL;
    }

    /* Convert wide string to UTF-8 */
    if (WideCharToMultiByte
        (CP_UTF8, 0, wide_str, -1, utf8_str, utf8_len, NULL, NULL) == 0) {
        free(wide_str);
        free(utf8_str);
        return NULL;
    }

    free(wide_str);
    return utf8_str;
#else
    /* On non-Windows, assume input is already UTF-8 or compatible */
    return strdup(ansi_str);
#endif
}

/* ============================================================================
 * Cell/Range Reference Functions
 * ============================================================================ */

void
lxw_parse_cell(const char *cell_str, lxw_row_t *row, lxw_col_t *col)
{
    if (row)
        *row = lxw_name_to_row(cell_str);
    if (col)
        *col = lxw_name_to_col(cell_str);
}

void
lxw_parse_cols(const char *cols_str, lxw_col_t *first_col,
               lxw_col_t *last_col)
{
    if (first_col)
        *first_col = lxw_name_to_col(cols_str);
    if (last_col)
        *last_col = lxw_name_to_col_2(cols_str);
}

void
lxw_parse_range(const char *range_str, lxw_row_t *first_row,
                lxw_col_t *first_col, lxw_row_t *last_row,
                lxw_col_t *last_col)
{
    if (first_row)
        *first_row = lxw_name_to_row(range_str);
    if (first_col)
        *first_col = lxw_name_to_col(range_str);
    if (last_row)
        *last_row = lxw_name_to_row_2(range_str);
    if (last_col)
        *last_col = lxw_name_to_col_2(range_str);
}

/* ============================================================================
 * Worksheet Write Functions with ANSI to UTF-8 conversion
 * ============================================================================ */

lxw_error
worksheet_write_string_lv(lxw_worksheet *worksheet, lxw_row_t row,
                          lxw_col_t col, const char *string,
                          lxw_format *format)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(string);

    err = worksheet_write_string(worksheet, row, col, utf8_str, format);
    free(utf8_str);
    return err;
}

lxw_error
worksheet_write_formula_lv(lxw_worksheet *worksheet, lxw_row_t row,
                           lxw_col_t col, const char *formula,
                           lxw_format *format)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(formula);

    err = worksheet_write_formula(worksheet, row, col, utf8_str, format);
    free(utf8_str);
    return err;
}

lxw_error
worksheet_write_url_lv(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                       const char *url, lxw_format *format)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(url);

    err = worksheet_write_url(worksheet, row, col, utf8_str, format);
    free(utf8_str);
    return err;
}

lxw_error
worksheet_write_comment_lv(lxw_worksheet *worksheet, lxw_row_t row,
                           lxw_col_t col, const char *string)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(string);

    err = worksheet_write_comment(worksheet, row, col, utf8_str);
    free(utf8_str);
    return err;
}

lxw_error
worksheet_set_header_lv(lxw_worksheet *worksheet, const char *header)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(header);

    err = worksheet_set_header(worksheet, utf8_str);
    free(utf8_str);
    return err;
}

lxw_error
worksheet_set_footer_lv(lxw_worksheet *worksheet, const char *footer)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(footer);

    err = worksheet_set_footer(worksheet, utf8_str);
    free(utf8_str);
    return err;
}

lxw_error
worksheet_merge_range_lv(lxw_worksheet *worksheet, lxw_row_t first_row,
                         lxw_col_t first_col, lxw_row_t last_row,
                         lxw_col_t last_col, const char *string,
                         lxw_format *format)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(string);

    err = worksheet_merge_range(worksheet, first_row, first_col, last_row,
                                last_col, utf8_str, format);
    free(utf8_str);
    return err;
}

/* ============================================================================
 * Chart Functions with ANSI to UTF-8 conversion
 * ============================================================================ */

lxw_chart_series *
chart_add_series_lv(lxw_chart *chart, const char *categories,
                    const char *values)
{
    lxw_chart_series *series;
    char *utf8_categories = ansi_to_utf8(categories);
    char *utf8_values = ansi_to_utf8(values);

    series = chart_add_series(chart, utf8_categories, utf8_values);
    free(utf8_categories);
    free(utf8_values);
    return series;
}

void
chart_series_set_name_lv(lxw_chart_series *series, const char *name)
{
    char *utf8_str = ansi_to_utf8(name);
    chart_series_set_name(series, utf8_str);
    free(utf8_str);
}

void
chart_axis_set_name_lv(lxw_chart_axis *axis, const char *name)
{
    char *utf8_str = ansi_to_utf8(name);
    chart_axis_set_name(axis, utf8_str);
    free(utf8_str);
}

void
chart_title_set_name_lv(lxw_chart *chart, const char *name)
{
    char *utf8_str = ansi_to_utf8(name);
    chart_title_set_name(chart, utf8_str);
    free(utf8_str);
}

/* ============================================================================
 * Format Functions with ANSI to UTF-8 conversion
 * ============================================================================ */

void
format_set_font_name_lv(lxw_format *format, const char *font_name)
{
    char *utf8_str = ansi_to_utf8(font_name);
    format_set_font_name(format, utf8_str);
    free(utf8_str);
}

void
format_set_num_format_lv(lxw_format *format, const char *num_format)
{
    char *utf8_str = ansi_to_utf8(num_format);
    format_set_num_format(format, utf8_str);
    free(utf8_str);
}

/* ============================================================================
 * Workbook Functions with ANSI to UTF-8 conversion (except filenames)
 * ============================================================================ */

lxw_worksheet *
workbook_add_worksheet_lv(lxw_workbook *workbook, const char *sheetname)
{
    lxw_worksheet *worksheet;
    char *utf8_str = ansi_to_utf8(sheetname);

    worksheet = workbook_add_worksheet(workbook, utf8_str);
    free(utf8_str);
    return worksheet;
}

lxw_chartsheet *
workbook_add_chartsheet_lv(lxw_workbook *workbook, const char *sheetname)
{
    lxw_chartsheet *chartsheet;
    char *utf8_str = ansi_to_utf8(sheetname);

    chartsheet = workbook_add_chartsheet(workbook, utf8_str);
    free(utf8_str);
    return chartsheet;
}

lxw_error
workbook_define_name_lv(lxw_workbook *workbook, const char *name,
                        const char *formula)
{
    lxw_error err;
    char *utf8_name = ansi_to_utf8(name);
    char *utf8_formula = ansi_to_utf8(formula);

    err = workbook_define_name(workbook, utf8_name, utf8_formula);
    free(utf8_name);
    free(utf8_formula);
    return err;
}

lxw_worksheet *
workbook_get_worksheet_by_name_lv(lxw_workbook *workbook, const char *name)
{
    lxw_worksheet *worksheet;
    char *utf8_str = ansi_to_utf8(name);

    worksheet = workbook_get_worksheet_by_name(workbook, utf8_str);
    free(utf8_str);
    return worksheet;
}

lxw_chartsheet *
workbook_get_chartsheet_by_name_lv(lxw_workbook *workbook, const char *name)
{
    lxw_chartsheet *chartsheet;
    char *utf8_str = ansi_to_utf8(name);

    chartsheet = workbook_get_chartsheet_by_name(workbook, utf8_str);
    free(utf8_str);
    return chartsheet;
}

lxw_error
workbook_validate_sheet_name_lv(lxw_workbook *workbook, const char *sheetname)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(sheetname);

    err = workbook_validate_sheet_name(workbook, utf8_str);
    free(utf8_str);
    return err;
}

lxw_error
workbook_set_custom_property_string_lv(lxw_workbook *workbook,
                                       const char *name, const char *value)
{
    lxw_error err;
    char *utf8_name = ansi_to_utf8(name);
    char *utf8_value = ansi_to_utf8(value);

    err =
        workbook_set_custom_property_string(workbook, utf8_name, utf8_value);
    free(utf8_name);
    free(utf8_value);
    return err;
}

void
worksheet_set_comments_author_lv(lxw_worksheet *worksheet, const char *author)
{
    char *utf8_str = ansi_to_utf8(author);
    worksheet_set_comments_author(worksheet, utf8_str);
    free(utf8_str);
}

lxw_error
chartsheet_set_header_lv(lxw_chartsheet *chartsheet, const char *header)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(header);

    err = chartsheet_set_header(chartsheet, utf8_str);
    free(utf8_str);
    return err;
}

lxw_error
chartsheet_set_footer_lv(lxw_chartsheet *chartsheet, const char *footer)
{
    lxw_error err;
    char *utf8_str = ansi_to_utf8(footer);

    err = chartsheet_set_footer(chartsheet, utf8_str);
    free(utf8_str);
    return err;
}

void
chart_series_set_trendline_name_lv(lxw_chart_series *series, const char *name)
{
    char *utf8_str = ansi_to_utf8(name);
    chart_series_set_trendline_name(series, utf8_str);
    free(utf8_str);
}

void
chart_axis_set_num_format_lv(lxw_chart_axis *axis, const char *num_format)
{
    char *utf8_str = ansi_to_utf8(num_format);
    chart_axis_set_num_format(axis, utf8_str);
    free(utf8_str);
}

void
chart_series_set_labels_num_format_lv(lxw_chart_series *series,
                                      const char *num_format)
{
    char *utf8_str = ansi_to_utf8(num_format);
    chart_series_set_labels_num_format(series, utf8_str);
    free(utf8_str);
}

/* ============================================================================
 * Functions with sheetname parameters
 * ============================================================================ */

void
chart_series_set_categories_lv(lxw_chart_series *series,
                               const char *sheetname, lxw_row_t first_row,
                               lxw_col_t first_col, lxw_row_t last_row,
                               lxw_col_t last_col)
{
    char *utf8_str = ansi_to_utf8(sheetname);
    chart_series_set_categories(series, utf8_str, first_row, first_col,
                                last_row, last_col);
    free(utf8_str);
}

void
chart_series_set_values_lv(lxw_chart_series *series, const char *sheetname,
                           lxw_row_t first_row, lxw_col_t first_col,
                           lxw_row_t last_row, lxw_col_t last_col)
{
    char *utf8_str = ansi_to_utf8(sheetname);
    chart_series_set_values(series, utf8_str, first_row, first_col, last_row,
                            last_col);
    free(utf8_str);
}

void
chart_series_set_name_range_lv(lxw_chart_series *series,
                               const char *sheetname, lxw_row_t row,
                               lxw_col_t col)
{
    char *utf8_str = ansi_to_utf8(sheetname);
    chart_series_set_name_range(series, utf8_str, row, col);
    free(utf8_str);
}

void
chart_axis_set_name_range_lv(lxw_chart_axis *axis, const char *sheetname,
                             lxw_row_t row, lxw_col_t col)
{
    char *utf8_str = ansi_to_utf8(sheetname);
    chart_axis_set_name_range(axis, utf8_str, row, col);
    free(utf8_str);
}

void
chart_title_set_name_range_lv(lxw_chart *chart, const char *sheetname,
                              lxw_row_t row, lxw_col_t col)
{
    char *utf8_str = ansi_to_utf8(sheetname);
    chart_title_set_name_range(chart, utf8_str, row, col);
    free(utf8_str);
}
