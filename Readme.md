# libxlsxwriter


Libxlsxwriter: A C library for creating Excel XLSX files.


![demo image](https://github.com/user-attachments/assets/3bb2a19f-76af-4d9b-8278-ce215c546381)



## The libxlsxwriter library

Libxlsxwriter is a C library that can be used to write text, numbers, formulas
and hyperlinks to multiple worksheets in an Excel 2007+ XLSX file.

It supports features such as:

- 100% compatible Excel XLSX files.
- Full Excel formatting.
- Merged cells.
- Defined names.
- Autofilters.
- Charts with secondary Y axis.
- Data validation and dropdown lists.
- Conditional formatting.
- Worksheet PNG/JPEG/GIF images.
- Cell comments.
- Support for adding Macros.
- Memory optimization mode for writing large files.
- Source code available on [GitHub](https://github.com/jmcnamara/libxlsxwriter).
- FreeBSD license.
- ANSI C.
- Works with GCC, Clang, Xcode, MSVC 2015, ICC, TCC, MinGW, MingGW-w64/32.
- Works on Linux, FreeBSD, OpenBSD, OS X, iOS and Windows. Also works on MSYS/MSYS2 and Cygwin.
- Compiles for 32 and 64 bit.
- Compiles and works on big and little endian systems.
- The only dependency is on `zlib`.

Here is an example that was used to create the spreadsheet shown above:


```C
#include "xlsxwriter.h"

int main() {

    lxw_workbook  *workbook  = workbook_new("test_secondary_axis.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "Data");
    lxw_chart     *chart     = workbook_add_chart(workbook, LXW_CHART_LINE);
    lxw_chart_series *series1, *series2;

    /* Write headers */
    worksheet_write_string(worksheet, 0, 0, "Month", NULL);
    worksheet_write_string(worksheet, 0, 1, "Temperature (°C)", NULL);
    worksheet_write_string(worksheet, 0, 2, "Revenue ($1000s)", NULL);

    /* Write some data with different scales to demonstrate secondary axis. */
    /* Column A: X-axis values (months) */
    const char *months[] = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug"};
    for (int i = 0; i < 8; i++) {
        worksheet_write_string(worksheet, i + 1, 0, months[i], NULL);
    }

    /* Column B: Small scale values - Temperature (primary Y-axis, left side) */
    double temperatures[] = {5.2, 7.8, 12.5, 16.3, 21.7, 25.4, 28.1, 26.8};
    for (int i = 0; i < 8; i++) {
        worksheet_write_number(worksheet, i + 1, 1, temperatures[i], NULL);
    }

    /* Column C: Large scale values - Revenue (secondary Y-axis, right side) */
    double revenue[] = {1250, 1580, 2340, 2890, 3420, 3950, 4200, 3880};
    for (int i = 0; i < 8; i++) {
        worksheet_write_number(worksheet, i + 1, 2, revenue[i], NULL);
    }

    /* Add series to primary Y-axis (left side) - Temperature */
    series1 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$9",
                                       "=Data!$B$2:$B$9",
                                       chart->y_axis);

    /* Set series name */
    chart_series_set_name(series1, "=Data!$B$1");

    /* Add series to secondary Y-axis (right side) - Revenue */
    series2 = chart_add_series_on_axis(chart,
                                       "=Data!$A$2:$A$9",
                                       "=Data!$C$2:$C$9",
                                       chart->y2_axis);

    /* Set series name */
    chart_series_set_name(series2, "=Data!$C$1");

    /* Configure chart title */
    chart_title_set_name(chart, "Monthly Temperature vs Revenue");

    /* Configure primary Y-axis (left side) */
    chart_axis_set_name(chart->y_axis, "Temperature (°C)");

    /* Configure secondary Y-axis (right side) */
    chart_axis_set_name(chart->y2_axis, "Revenue ($1000s)");

    /* Position the secondary Y-axis title on the right side of the chart */
    lxw_chart_layout layout = {
        .x = 0.686,
        .y = 0.314,
        .width = 0.0,
        .height = 0.0,
        .has_inner = LXW_FALSE
    };
    chart_axis_set_name_layout(chart->y2_axis, &layout);

    /* Configure X-axis */
    chart_axis_set_name(chart->x_axis, "Month");


    /* Insert the chart into the worksheet */
    worksheet_insert_chart(worksheet, CELL("E2"), chart);


    /* Save the workbook */
    lxw_error result = workbook_close(workbook);

    if (result == LXW_NO_ERROR) {
        printf("\n✓ SUCCESS! Created test_secondary_axis.xlsx\n\n");
        return 0;
    } else {
        printf("\n✗ ERROR creating workbook: %d\n", result);
        return 1;
    }
}
```


See the [full documentation](http://libxlsxwriter.github.io) for the getting
started guide, a tutorial, the main API documentation and examples.

