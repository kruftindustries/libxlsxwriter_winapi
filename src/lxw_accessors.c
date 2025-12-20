/*
 * lxw_accessors.c - Accessor functions for LabVIEW compatibility
 *
 * These functions provide access to chart axis pointers which are
 * needed for chart_add_series_on_axis() and axis configuration.
 */

#include "xlsxwriter.h"

/*
 * Get the primary X-axis from a chart.
 */
lxw_chart_axis* chart_get_x_axis(lxw_chart *chart) {
    if (chart)
        return chart->x_axis;
    return NULL;
}

/*
 * Get the primary Y-axis from a chart.
 */
lxw_chart_axis* chart_get_y_axis(lxw_chart *chart) {
    if (chart)
        return chart->y_axis;
    return NULL;
}

/*
 * Get the secondary Y-axis from a chart.
 */
lxw_chart_axis* chart_get_y2_axis(lxw_chart *chart) {
    if (chart)
        return chart->y2_axis;
    return NULL;
}
