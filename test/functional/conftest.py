"""
Pytest configuration for functional tests.
Adds timing measurements and displays the slowest test at the end.
"""

import pytest
import time


# Store test durations
test_durations = []


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    """Hook to capture test duration after each test."""
    outcome = yield
    report = outcome.get_result()

    if report.when == "call":
        # Store test name and duration
        test_durations.append({
            'name': item.nodeid,
            'duration': report.duration
        })


def pytest_terminal_summary(terminalreporter, exitstatus, config):
    """Hook to display the slowest test at the end of the test run."""
    if test_durations:
        # Find the slowest test
        slowest = max(test_durations, key=lambda x: x['duration'])

        terminalreporter.write_sep("=", "Slowest Test")
        terminalreporter.write_line(
            f"  {slowest['name']}: {slowest['duration']:.4f}s"
        )
        terminalreporter.write_line("")
