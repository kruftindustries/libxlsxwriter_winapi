/*
 * Test program to verify area chart secondary Y-axis functionality.
 * Tests area chart with two series on different Y-axes.
 */

#include "xlsxwriter.h"
#include <stdio.h>

int main() {
    printf("=== Area Chart Secondary Y-Axis Test ===\n\n");

    lxw_workbook  *workbook  = workbook_new("test_area_y2.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "Data");
    lxw_chart     *chart     = workbook_add_chart(workbook, LXW_CHART_AREA);
    lxw_chart_series *series1, *series2;

    /* Write headers */
    worksheet_write_string(worksheet, 0, 0, "Quarter", NULL);
    worksheet_write_string(worksheet, 0, 1, "Units Sold (1000s)", NULL);
    worksheet_write_string(worksheet, 0, 2, "Market Share (%)", NULL);

    /* Write some data with different scales to demonstrate secondary axis. */
    /* Column A: X-axis values (quarters) */
    const char *quarters[] = {"Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8"};
    for (int i = 0; i < 8; i++) {
        worksheet_write_string(worksheet, i + 1, 0, quarters[i], NULL);
    }

    /* Column B: Small scale values - Units Sold (primary Y-axis, left side) */
    double units[] = {12.5, 15.3, 18.7, 22.1, 25.8, 28.3, 30.2, 31.5};
    for (int i = 0; i < 8; i++) {
        worksheet_write_number(worksheet, i + 1, 1, units[i], NULL);
    }

    /* Column C: Different scale values - Market Share % (secondary Y-axis, right side) */
    double market_share[] = {8.5, 12.3, 15.7, 18.2, 22.5, 26.8, 29.3, 32.1};
    for (int i = 0; i < 8; i++) {
        worksheet_write_number(worksheet, i + 1, 2, market_share[i], NULL);
    }

    printf("1. Writing test data (8 quarters, 2 series with different scales)...\n");

    /* Add series to primary Y-axis (left side) - Units Sold */
    series1 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$9",
                                       "=Data!$B$2:$B$9",
                                       chart->y_axis);

    /* Set series name */
    chart_series_set_name(series1, "=Data!$B$1");

    printf("2. Added primary series (Units Sold) to left Y-axis...\n");

    /* Add series to secondary Y-axis (right side) - Market Share */
    series2 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$9",
                                       "=Data!$C$2:$C$9",
                                       chart->y2_axis);

    /* Set series name */
    chart_series_set_name(series2, "=Data!$C$1");

    printf("3. Added secondary series (Market Share) to right Y-axis...\n");

    /* Configure chart title */
    chart_title_set_name(chart, "Units Sold and Market Share by Quarter");

    /* Configure X-axis */
    chart_axis_set_name(chart->x_axis, "Quarter");

    /* Configure primary Y-axis (left side) */
    chart_axis_set_name(chart->y_axis, "Units Sold (1000s)");

    /* Configure secondary Y-axis (right side) */
    chart_axis_set_name(chart->y2_axis, "Market Share (%)");

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
        printf("\n✓ SUCCESS! Created test_area_y2.xlsx\n\n");
        printf("Expected behavior:\n");
        printf("  - Area chart with TWO Y-axes\n");
        printf("  - Units Sold series on left Y-axis (0-35 range)\n");
        printf("  - Market Share series on right Y-axis (0-35%% range)\n");
        printf("  - Both series plotted against quarters on X-axis\n");
        printf("  - Two areaChart elements in the XML\n");
        printf("  - Four axes total (2 X-axes, 2 Y-axes)\n\n");
        printf("Open test_area_y2.xlsx in Excel to verify!\n");
        return 0;
    } else {
        printf("\n✗ ERROR creating workbook: %d\n", result);
        return 1;
    }
}
