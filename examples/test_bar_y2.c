/*
 * Test program to verify bar chart secondary Y-axis functionality.
 * Tests horizontal bar chart with two series on different Y-axes.
 */

#include "xlsxwriter.h"
#include <stdio.h>

int main() {
    printf("=== Bar Chart Secondary Y-Axis Test ===\n\n");

    lxw_workbook  *workbook  = workbook_new("test_bar_y2.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "Data");
    lxw_chart     *chart     = workbook_add_chart(workbook, LXW_CHART_BAR);
    lxw_chart_series *series1, *series2;

    /* Write headers */
    worksheet_write_string(worksheet, 0, 0, "Department", NULL);
    worksheet_write_string(worksheet, 0, 1, "Employees", NULL);
    worksheet_write_string(worksheet, 0, 2, "Budget ($M)", NULL);

    /* Write some data with different scales to demonstrate secondary axis. */
    /* Column A: X-axis values (departments) */
    const char *departments[] = {"Engineering", "Sales", "Marketing", "Operations", "HR", "Finance"};
    for (int i = 0; i < 6; i++) {
        worksheet_write_string(worksheet, i + 1, 0, departments[i], NULL);
    }

    /* Column B: Small scale values - Employees (primary Y-axis, left side) */
    double employees[] = {45, 32, 18, 28, 12, 15};
    for (int i = 0; i < 6; i++) {
        worksheet_write_number(worksheet, i + 1, 1, employees[i], NULL);
    }

    /* Column C: Large scale values - Budget in millions (secondary Y-axis, right side) */
    double budget[] = {8.5, 6.2, 3.8, 5.1, 2.2, 2.8};
    for (int i = 0; i < 6; i++) {
        worksheet_write_number(worksheet, i + 1, 2, budget[i], NULL);
    }

    printf("1. Writing test data (6 departments, 2 series with different scales)...\n");

    /* Add series to primary Y-axis (left side) - Employees */
    series1 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$7",
                                       "=Data!$B$2:$B$7",
                                       chart->y_axis);

    /* Set series name */
    chart_series_set_name(series1, "=Data!$B$1");

    printf("2. Added primary series (Employees) to left Y-axis...\n");

    /* Add series to secondary Y-axis (right side) - Budget */
    series2 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$7",
                                       "=Data!$C$2:$C$7",
                                       chart->y2_axis);

    /* Set series name */
    chart_series_set_name(series2, "=Data!$C$1");

    printf("3. Added secondary series (Budget) to right Y-axis...\n");

    /* Configure chart title */
    chart_title_set_name(chart, "Department Employees and Budget");

    /* Configure X-axis */
    chart_axis_set_name(chart->x_axis, "Department");

    /* Configure primary Y-axis (left side) */
    chart_axis_set_name(chart->y_axis, "Employees");

    /* Configure secondary Y-axis (right side) */
    chart_axis_set_name(chart->y2_axis, "Budget ($M)");

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
        printf("\n✓ SUCCESS! Created test_bar_y2.xlsx\n\n");
        printf("Expected behavior:\n");
        printf("  - Horizontal bar chart with TWO Y-axes\n");
        printf("  - Employees series on left Y-axis (0-50 range)\n");
        printf("  - Budget series on right Y-axis (0-10M range)\n");
        printf("  - Both series plotted against departments on X-axis\n");
        printf("  - Two barChart elements in the XML\n");
        printf("  - Four axes total (2 X-axes, 2 Y-axes)\n\n");
        printf("Open test_bar_y2.xlsx in Excel to verify!\n");
        return 0;
    } else {
        printf("\n✗ ERROR creating workbook: %d\n", result);
        return 1;
    }
}
