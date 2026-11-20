"""Shared pytest setup.

On Windows, Qt objects that are still alive when the interpreter shuts down
can crash during teardown *after* every test has passed, turning a green run
into exit code 1. Once the session is finished and the result is known, exit
immediately with pytest's own status so teardown crashes cannot mask it.
"""
import os
import sys


def pytest_sessionfinish(session, exitstatus):
    if sys.platform == "win32" and os.environ.get("PORTABLEEXCEL_HARD_EXIT", "1") == "1":
        sys.stdout.flush()
        sys.stderr.flush()
        os._exit(int(exitstatus))
