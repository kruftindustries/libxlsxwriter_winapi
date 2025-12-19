/*
 * Test program to verify scatter chart secondary Y-axis functionality.
 * Tests scatter chart with two series on different Y-axes.
 */

#include "xlsxwriter.h"
#include <stdio.h>

int main() {
    printf("=== Scatter Chart Secondary Y-Axis Test ===\n\n");

    lxw_workbook  *workbook  = workbook_new("test_scatter_y2.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "Data");
    lxw_chart     *chart     = workbook_add_chart(workbook, LXW_CHART_SCATTER);
    lxw_chart_series *series1, *series2;

    /* Write headers */
    worksheet_write_string(worksheet, 0, 0, "Time (hours)", NULL);
    worksheet_write_string(worksheet, 0, 1, "Velocity (m/s)", NULL);
    worksheet_write_string(worksheet, 0, 2, "Distance (km)", NULL);

    /* Write some data with different scales to demonstrate secondary axis. */
    /* Column A: X-axis values (time in hours) */
    double time[] = {0, 1, 2, 3, 4, 5, 6, 7, 8};
    for (int i = 0; i < 9; i++) {
        worksheet_write_number(worksheet, i + 1, 0, time[i], NULL);
    }

    /* Column B: Small scale values - Velocity in m/s (primary Y-axis, left side) */
    double velocity[] = {0, 5.5, 12.3, 18.7, 22.1, 23.8, 24.2, 24.5, 24.7};
    for (int i = 0; i < 9; i++) {
        worksheet_write_number(worksheet, i + 1, 1, velocity[i], NULL);
    }

    /* Column C: Large scale values - Distance in km (secondary Y-axis, right side) */
    double distance[] = {0, 10, 35, 75, 130, 200, 285, 385, 500};
    for (int i = 0; i < 9; i++) {
        worksheet_write_number(worksheet, i + 1, 2, distance[i], NULL);
    }

    printf("1. Writing test data (9 time points, 2 series with different scales)...\n");

    /* Add series to primary Y-axis (left side) - Velocity */
    series1 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$10",
                                       "=Data!$B$2:$B$10",
                                       chart->y_axis);

    /* Set series name */
    chart_series_set_name(series1, "=Data!$B$1");

    printf("2. Added primary series (Velocity) to left Y-axis...\n");

    /* Add series to secondary Y-axis (right side) - Distance */
    series2 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$10",
                                       "=Data!$C$2:$C$10",
                                       chart->y2_axis);

    /* Set series name */
    chart_series_set_name(series2, "=Data!$C$1");

    printf("3. Added secondary series (Distance) to right Y-axis...\n");

    /* Configure chart title */
    chart_title_set_name(chart, "Velocity and Distance vs Time");

    /* Configure X-axis */
    chart_axis_set_name(chart->x_axis, "Time (hours)");

    /* Configure primary Y-axis (left side) */
    chart_axis_set_name(chart->y_axis, "Velocity (m/s)");

    /* Configure secondary Y-axis (right side) */
    chart_axis_set_name(chart->y2_axis, "Distance (km)");

    /* Position the secondary Y-axis title on the right side of the chart */
    lxw_chart_layout layout = {
        .x = 0.686,
        .y = 0.314,
        .width = 0.0,
        .height = 0.0,
        .has_inner = LXW_FALSE
    };
    chart_axis_set_name_layout(chart->y2_axis, &layout);

    printf("4. Configured chart title and axis labels...\n");

    /* Insert the chart into the worksheet */
    worksheet_insert_chart(worksheet, CELL("E2"), chart);

    printf("5. Inserted chart into worksheet...\n");

    /* Save the workbook */
    printf("6. Saving workbook...\n");
    lxw_error result = workbook_close(workbook);

    if (result == LXW_NO_ERROR) {
        printf("\n✓ SUCCESS! Created test_scatter_y2.xlsx\n\n");
        printf("Expected behavior:\n");
        printf("  - Scatter chart with TWO Y-axes\n");
        printf("  - Velocity series on left Y-axis (0-25 m/s range)\n");
        printf("  - Distance series on right Y-axis (0-500 km range)\n");
        printf("  - Both series plotted against time on X-axis\n");
        printf("  - Two scatterChart elements in the XML\n");
        printf("  - Four axes total (2 X-axes, 2 Y-axes)\n\n");
        printf("Open test_scatter_y2.xlsx in Excel to verify!\n");
        return 0;
    } else {
        printf("\n✗ ERROR creating workbook: %d\n", result);
        return 1;
    }
}
