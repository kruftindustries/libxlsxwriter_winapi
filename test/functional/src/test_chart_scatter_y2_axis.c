/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test scatter chart with secondary Y-axis (y2_axis) functionality.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "xlsxwriter.h"

int main() {

    lxw_workbook  *workbook  = workbook_new("test_chart_scatter_y2_axis.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);
    lxw_chart     *chart     = workbook_add_chart(workbook, LXW_CHART_SCATTER);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 40262272;
    chart->axis_id_2 = 40260352;
    chart->axis_id_3 = 40263296;
    chart->axis_id_4 = 40264320;

    uint8_t data[5][3] = {
        {1, 2,  10},
        {2, 4,  40},
        {3, 6,  90},
        {4, 8,  160},
        {5, 10, 250}
    };

    int row, col;
    for (row = 0; row < 5; row++)
        for (col = 0; col < 3; col++)
            worksheet_write_number(worksheet, row, col, data[row][col], NULL);

    /* Add series on primary Y-axis (left) */
    chart_add_series(chart,
         "=Sheet1!$A$1:$A$5",
         "=Sheet1!$B$1:$B$5"
    );

    /* Add series on secondary Y-axis (right) */
    chart_add_series_on_axis(chart,
         "=Sheet1!$A$1:$A$5",
         "=Sheet1!$C$1:$C$5",
         chart->y2_axis
    );

    /* Set axis names to verify they appear on correct sides */
    chart_axis_set_name(chart->y_axis, "Primary Y (left)");
    chart_axis_set_name(chart->y2_axis, "Secondary Y (right)");

    worksheet_insert_chart(worksheet, CELL("E9"), chart);

    return workbook_close(workbook);
}
