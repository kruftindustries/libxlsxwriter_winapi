#!/bin/bash
# Verification script for scatter chart secondary Y-axis functionality.
# This script compiles the test, runs it, and verifies the XML output.

set -e

echo "=== Secondary Y-Axis Verification Test ==="
echo ""

# Clean up any previous test
rm -f test_scatter_y2 test_chart_scatter_y2_axis.xlsx
rm -rf test_y2_verify

# Compile the test
echo "1. Compiling test program..."
gcc -o test_scatter_y2 test/functional/src/test_chart_scatter_y2_axis.c \
    -I./include -L./lib -lxlsxwriter -lz

# Run the test
echo "2. Running test program..."
LD_LIBRARY_PATH=./lib:$LD_LIBRARY_PATH ./test_scatter_y2

# Extract the xlsx file
echo "3. Extracting xlsx file..."
mkdir test_y2_verify
unzip -q test_chart_scatter_y2_axis.xlsx -d test_y2_verify

# Verify the chart XML
CHART_XML="test_y2_verify/xl/charts/chart1.xml"
echo "4. Verifying chart XML..."

# Check for secondary Y-axis on right position
if grep -q '<c:axPos val="r"/>' "$CHART_XML"; then
    echo "   ✓ Secondary Y-axis position: RIGHT"
else
    echo "   ✗ FAIL: Secondary Y-axis not on right"
    exit 1
fi

# Check for crosses at max
if grep -q '<c:crosses val="max"/>' "$CHART_XML"; then
    echo "   ✓ Secondary Y-axis crosses at MAX"
else
    echo "   ✗ FAIL: Secondary Y-axis doesn't cross at max"
    exit 1
fi

# Check for high tick label position on Y2 axis
if grep -q '<c:tickLblPos val="high"/>' "$CHART_XML"; then
    echo "   ✓ Secondary Y-axis tick labels: HIGH"
else
    echo "   ✗ FAIL: Secondary Y-axis tick labels not high"
    exit 1
fi

# Check for low tick label position on Y1 axis
if grep -q '<c:tickLblPos val="low"/>' "$CHART_XML"; then
    echo "   ✓ Primary Y-axis tick labels: LOW"
else
    echo "   ✗ FAIL: Primary Y-axis tick labels not low"
    exit 1
fi

# Check for axis titles
if grep -q 'Primary Y (left)' "$CHART_XML"; then
    echo "   ✓ Primary Y-axis title present"
else
    echo "   ✗ FAIL: Primary Y-axis title missing"
    exit 1
fi

if grep -q 'Secondary Y (right)' "$CHART_XML"; then
    echo "   ✓ Secondary Y-axis title present"
else
    echo "   ✗ FAIL: Secondary Y-axis title missing"
    exit 1
fi

echo ""
echo "=== All verification checks passed! ==="

# Clean up
rm -f test_scatter_y2 test_chart_scatter_y2_axis.xlsx
rm -rf test_y2_verify
