"""Shared pytest setup.

On Windows, Qt can crash while its DLLs are unloaded at process exit, *after*
every test has passed, which turns a green run into exit code 1. Once the
session is finished and the result is known, end the process with
TerminateProcess (which skips DLL teardown) and pytest's own exit status.
"""
import os
import sys


def pytest_sessionfinish(session, exitstatus):
    if sys.platform != "win32" or os.environ.get("PORTABLEEXCEL_HARD_EXIT", "1") != "1":
        return
    sys.stdout.flush()
    sys.stderr.flush()
    import ctypes
    k32 = ctypes.windll.kernel32
    k32.TerminateProcess(k32.GetCurrentProcess(), int(exitstatus))
