/*
 * Test program to create a styled scatter chart in a chartsheet.
 * This is the chartsheet version of test_scatter_styled.c
 * Features:
 * - Circle markers with custom colors (blue/orange to approximate accent1/accent2)
 * - No lines connecting markers
 * - Legend at bottom
 * - Secondary Y-axis
 * - Formatted axis titles
 * - Chart displayed in a dedicated chartsheet
 */

#include "xlsxwriter.h"
#include <stdio.h>

int main() {
    printf("=== Styled Scatter Chart in Chartsheet ===\n\n");

    lxw_workbook   *workbook   = workbook_new("test_scatter_chartsheet.xlsx");
    lxw_worksheet  *worksheet  = workbook_add_worksheet(workbook, "Data");
    lxw_chartsheet *chartsheet = workbook_add_chartsheet(workbook, "Chart");
    lxw_chart      *chart      = workbook_add_chart(workbook, LXW_CHART_SCATTER);
    lxw_chart_series *series1, *series2;

    /* Write headers */
    worksheet_write_string(worksheet, 0, 0, "Time (hours)", NULL);
    worksheet_write_string(worksheet, 0, 1, "Velocity (m/s)", NULL);
    worksheet_write_string(worksheet, 0, 2, "Distance (km)", NULL);

    /* Write time data (X-axis) */
    double time[] = {0, 1, 2, 3, 4, 5, 6, 7, 8};
    for (int i = 0; i < 9; i++) {
        worksheet_write_number(worksheet, i + 1, 0, time[i], NULL);
    }

    /* Write velocity data (primary Y-axis, left side) */
    double velocity[] = {0, 5.5, 12.3, 18.7, 22.1, 23.8, 24.2, 24.5, 24.7};
    for (int i = 0; i < 9; i++) {
        worksheet_write_number(worksheet, i + 1, 1, velocity[i], NULL);
    }

    /* Write distance data (secondary Y-axis, right side) */
    double distance[] = {0, 5, 25, 70, 135, 215, 285, 385, 500};
    for (int i = 0; i < 9; i++) {
        worksheet_write_number(worksheet, i + 1, 2, distance[i], NULL);
    }

    printf("1. Writing test data (9 time points, 2 series)...\n");

    /* Configure colors to approximate Excel's accent1 (blue) and accent2 (orange) */
    lxw_color_t accent1_color = 0x4472C4;  /* Excel default accent1: blue */
    lxw_color_t accent2_color = 0xED7D31;  /* Excel default accent2: orange */

    /* Configure series 1 (Velocity) on primary Y-axis */
    series1 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$10",
                                       "=Data!$B$2:$B$10",
                                       chart->y_axis);

    chart_series_set_name(series1, "=Data!$B$1");

    /* Set marker style for series 1 - circle with blue fill */
    chart_series_set_marker_type(series1, LXW_CHART_MARKER_CIRCLE);
    chart_series_set_marker_size(series1, 5);

    lxw_chart_fill marker1_fill = {.color = accent1_color};
    chart_series_set_marker_fill(series1, &marker1_fill);

    lxw_chart_line marker1_line = {.color = accent1_color};
    chart_series_set_marker_line(series1, &marker1_line);

    /* Remove connecting line for series 1 (markers only) */
    lxw_chart_line no_line = {.none = 1};
    chart_series_set_line(series1, &no_line);

    printf("2. Added primary series (Velocity) with blue circle markers...\n");

    /* Configure series 2 (Distance) on secondary Y-axis */
    series2 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$10",
                                       "=Data!$C$2:$C$10",
                                       chart->y2_axis);

    chart_series_set_name(series2, "=Data!$C$1");

    /* Set marker style for series 2 - circle with orange fill */
    chart_series_set_marker_type(series2, LXW_CHART_MARKER_CIRCLE);
    chart_series_set_marker_size(series2, 5);

    lxw_chart_fill marker2_fill = {.color = accent2_color};
    chart_series_set_marker_fill(series2, &marker2_fill);

    lxw_chart_line marker2_line = {.color = accent2_color};
    chart_series_set_marker_line(series2, &marker2_line);

    /* Remove connecting line for series 2 (markers only) */
    chart_series_set_line(series2, &no_line);

    printf("3. Added secondary series (Distance) with orange circle markers...\n");

    /* Configure chart title */
    chart_title_set_name(chart, "Velocity (m/s) vs Distance (km)");

    /* Configure X-axis */
    chart_axis_set_name(chart->x_axis, "Time (hours)");

    /* Configure primary Y-axis (left side) */
    chart_axis_set_name(chart->y_axis, "Velocity (m/s)");

    /* Configure secondary Y-axis (right side) */
    chart_axis_set_name(chart->y2_axis, "Distance (km)");

    /* Position legend at bottom */
    chart_legend_set_position(chart, LXW_CHART_LEGEND_BOTTOM);

    printf("4. Configured chart title, axis labels, and legend...\n");

    /* Add the chart to the chartsheet (instead of inserting into worksheet) */
    chartsheet_set_chart(chartsheet, chart);

    /* Make the chartsheet the active sheet when the workbook is opened */
    chartsheet_activate(chartsheet);

    printf("5. Added chart to chartsheet...\n");

    /* Save the workbook */
    printf("6. Saving workbook...\n");
    lxw_error result = workbook_close(workbook);

    if (result == LXW_NO_ERROR) {
        printf("\n✓ SUCCESS! Created test_scatter_chartsheet.xlsx\n\n");
        printf("Styling features:\n");
        printf("  - Circle markers (size 5) for both series\n");
        printf("  - Series 1: Blue markers (accent1 approximation)\n");
        printf("  - Series 2: Orange markers (accent2 approximation)\n");
        printf("  - No connecting lines (markers only)\n");
        printf("  - Legend positioned at bottom\n");
        printf("  - Two Y-axes (left for Velocity, right for Distance)\n");
        printf("  - Chart in dedicated chartsheet (not embedded in worksheet)\n\n");
        printf("Open test_scatter_chartsheet.xlsx in Excel to verify!\n");
        return 0;
    } else {
        printf("\n✗ ERROR creating workbook: %d\n", result);
        return 1;
    }
}
