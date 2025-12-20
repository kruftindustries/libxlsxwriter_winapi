/*
 * Tests for the libxlsxwriter library.
 *
 * Test cases for secondary Y-axis (y2_axis) functionality.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/xlsxwriter/chart.h"

/*
 * Test that the secondary Y-axis has the correct default position (right).
 */
CTEST(chart_secondary_axis, y2_axis_position_default) {
    lxw_chart *chart = lxw_chart_new(LXW_CHART_SCATTER);

    ASSERT_EQUAL(LXW_CHART_AXIS_RIGHT, chart->y2_axis->axis_position);

    lxw_chart_free(chart);
}

/*
 * Test that the primary Y-axis has default label position LOW.
 */
CTEST(chart_secondary_axis, y_axis_label_position_default) {
    lxw_chart *chart = lxw_chart_new(LXW_CHART_SCATTER);

    ASSERT_EQUAL(LXW_CHART_AXIS_LABEL_POSITION_LOW, chart->y_axis->label_position);

    lxw_chart_free(chart);
}

/*
 * Test that the secondary Y-axis has default label position HIGH.
 */
CTEST(chart_secondary_axis, y2_axis_label_position_default) {
    lxw_chart *chart = lxw_chart_new(LXW_CHART_SCATTER);

    ASSERT_EQUAL(LXW_CHART_AXIS_LABEL_POSITION_HIGH, chart->y2_axis->label_position);

    lxw_chart_free(chart);
}

/*
 * Test scatter chart with secondary Y-axis generates correct axis configuration.
 * The secondary Y-axis should have:
 * - axPos val="r" (right position)
 * - crosses val="max" (crosses at max of X-axis)
 * - tickLblPos val="high"
 * The primary Y-axis should have:
 * - tickLblPos val="low"
 */
CTEST(chart_secondary_axis, scatter_y2_axis_xml) {

    char* got;

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chart *chart = lxw_chart_new(LXW_CHART_SCATTER);
    chart->file = testfile;

    /* Add series on primary Y-axis */
    lxw_chart_series *series1 = chart_add_series(chart,
        "Sheet1!$A$1:$A$3", "Sheet1!$B$1:$B$3");

    /* Add series on secondary Y-axis */
    lxw_chart_series *series2 = chart_add_series_on_axis(chart,
        "Sheet1!$A$1:$A$3", "Sheet1!$C$1:$C$3", chart->y2_axis);

    (void)series1;  /* Suppress unused variable warning */
    (void)series2;

    lxw_chart_assemble_xml_file(chart);

    /* Read the generated XML */
    fflush(testfile);
    int file_size = ftell(testfile);
    got = (char*)calloc(file_size + 1, 1);
    rewind(testfile);
    (void)fread(got, file_size, 1, testfile);

    /* Verify key Y2 axis elements are present */
    ASSERT_NOT_NULL(strstr(got, "<c:axPos val=\"r\"/>"));
    ASSERT_NOT_NULL(strstr(got, "<c:crosses val=\"max\"/>"));
    ASSERT_NOT_NULL(strstr(got, "<c:tickLblPos val=\"high\"/>"));

    /* Verify Y1 axis has low tick label position */
    ASSERT_NOT_NULL(strstr(got, "<c:tickLblPos val=\"low\"/>"));

    free(got);
    fclose(testfile);
    lxw_chart_free(chart);
}
