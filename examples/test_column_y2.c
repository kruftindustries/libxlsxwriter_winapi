/*
 * Test program to verify column chart secondary Y-axis functionality.
 * Tests vertical column chart with two series on different Y-axes.
 */

#include "xlsxwriter.h"
#include <stdio.h>

int main() {
    printf("=== Column Chart Secondary Y-Axis Test ===\n\n");

    lxw_workbook  *workbook  = workbook_new("test_column_y2.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "Data");
    lxw_chart     *chart     = workbook_add_chart(workbook, LXW_CHART_COLUMN);
    lxw_chart_series *series1, *series2;

    /* Write headers */
    worksheet_write_string(worksheet, 0, 0, "Product", NULL);
    worksheet_write_string(worksheet, 0, 1, "Sales (Units)", NULL);
    worksheet_write_string(worksheet, 0, 2, "Profit ($K)", NULL);

    /* Write some data with different scales to demonstrate secondary axis. */
    /* Column A: X-axis values (products) */
    const char *products[] = {"Widget A", "Widget B", "Widget C", "Widget D", "Widget E", "Widget F"};
    for (int i = 0; i < 6; i++) {
        worksheet_write_string(worksheet, i + 1, 0, products[i], NULL);
    }

    /* Column B: Large scale values - Sales Units (primary Y-axis, left side) */
    double sales[] = {1250, 1850, 2340, 1920, 2680, 3120};
    for (int i = 0; i < 6; i++) {
        worksheet_write_number(worksheet, i + 1, 1, sales[i], NULL);
    }

    /* Column C: Small scale values - Profit in thousands (secondary Y-axis, right side) */
    double profit[] = {45.2, 68.5, 92.3, 71.8, 105.6, 128.4};
    for (int i = 0; i < 6; i++) {
        worksheet_write_number(worksheet, i + 1, 2, profit[i], NULL);
    }

    printf("1. Writing test data (6 products, 2 series with different scales)...\n");

    /* Add series to primary Y-axis (left side) - Sales */
    series1 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$7",
                                       "=Data!$B$2:$B$7",
                                       chart->y_axis);

    /* Set series name */
    chart_series_set_name(series1, "=Data!$B$1");

    printf("2. Added primary series (Sales) to left Y-axis...\n");

    /* Add series to secondary Y-axis (right side) - Profit */
    series2 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$7",
                                       "=Data!$C$2:$C$7",
                                       chart->y2_axis);

    /* Set series name */
    chart_series_set_name(series2, "=Data!$C$1");

    printf("3. Added secondary series (Profit) to right Y-axis...\n");

    /* Configure chart title */
    chart_title_set_name(chart, "Product Sales and Profit");

    /* Configure X-axis */
    chart_axis_set_name(chart->x_axis, "Product");

    /* Configure primary Y-axis (left side) */
    chart_axis_set_name(chart->y_axis, "Sales (Units)");

    /* Configure secondary Y-axis (right side) */
    chart_axis_set_name(chart->y2_axis, "Profit ($K)");

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
        printf("\n✓ SUCCESS! Created test_column_y2.xlsx\n\n");
        printf("Expected behavior:\n");
        printf("  - Column chart with TWO Y-axes\n");
        printf("  - Sales series on left Y-axis (0-3500 range)\n");
        printf("  - Profit series on right Y-axis (0-140K range)\n");
        printf("  - Both series plotted against products on X-axis\n");
        printf("  - Two barChart elements in the XML (with barDir=\"col\")\n");
        printf("  - Four axes total (2 X-axes, 2 Y-axes)\n\n");
        printf("Open test_column_y2.xlsx in Excel to verify!\n");
        return 0;
    } else {
        printf("\n✗ ERROR creating workbook: %d\n", result);
        return 1;
    }
}
