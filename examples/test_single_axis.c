/*
 * Test program to verify single Y-axis functionality (baseline test).
 * This ensures our secondary axis changes didn't break normal charts.
 */

#include "xlsxwriter.h"
#include <stdio.h>

int main() {
    printf("=== Single Y-Axis Test (Baseline) ===\n\n");

    lxw_workbook  *workbook  = workbook_new("test_single_axis.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "Sales");
    lxw_chart     *chart     = workbook_add_chart(workbook, LXW_CHART_LINE);

    /* Write headers */
    worksheet_write_string(worksheet, 0, 0, "Quarter", NULL);
    worksheet_write_string(worksheet, 0, 1, "Sales", NULL);
    worksheet_write_string(worksheet, 0, 2, "Profit", NULL);

    printf("1. Writing test data...\n");

    /* Write quarter labels */
    const char *quarters[] = {"Q1", "Q2", "Q3", "Q4"};
    for (int i = 0; i < 4; i++) {
        worksheet_write_string(worksheet, i + 1, 0, quarters[i], NULL);
    }

    /* Write sales data */
    double sales[] = {120, 180, 220, 195};
    for (int i = 0; i < 4; i++) {
        worksheet_write_number(worksheet, i + 1, 1, sales[i], NULL);
    }

    /* Write profit data */
    double profit[] = {45, 68, 82, 71};
    for (int i = 0; i < 4; i++) {
        worksheet_write_number(worksheet, i + 1, 2, profit[i], NULL);
    }

    printf("2. Adding two series to PRIMARY Y-axis only...\n");

    /* Add first series - Sales (using default/primary axis) */
    lxw_chart_series *series1 = chart_add_series(chart,
                                                  "=Sales!$A$2:$A$5",
                                                  "=Sales!$B$2:$B$5");
    chart_series_set_name(series1, "=Sales!$B$1");

    /* Add second series - Profit (using default/primary axis) */
    lxw_chart_series *series2 = chart_add_series(chart,
                                                  "=Sales!$A$2:$A$5",
                                                  "=Sales!$C$2:$C$5");
    chart_series_set_name(series2, "=Sales!$C$1");

    printf("3. Configuring chart...\n");

    /* Configure chart */
    chart_title_set_name(chart, "Quarterly Sales and Profit");
    chart_axis_set_name(chart->x_axis, "Quarter");
    chart_axis_set_name(chart->y_axis, "Amount ($K)");

    /* Insert chart */
    worksheet_insert_chart(worksheet, CELL("E2"), chart);

    printf("4. Saving workbook...\n");

    lxw_error result = workbook_close(workbook);

    if (result == LXW_NO_ERROR) {
        printf("\n✓ SUCCESS! Created test_single_axis.xlsx\n\n");
        printf("Expected behavior:\n");
        printf("  - Chart with ONE Y-axis (standard chart)\n");
        printf("  - Two series (Sales and Profit) on same scale\n");
        printf("  - Chart style and color files should still be generated\n\n");
        printf("This verifies our changes didn't break normal charts!\n");
        return 0;
    } else {
        printf("\n✗ ERROR creating workbook: %d\n", result);
        return 1;
    }
}
