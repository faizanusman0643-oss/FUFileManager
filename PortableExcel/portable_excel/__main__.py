"""Entry point: python -m portable_excel [file]"""
import logging
import sys


def _exit(rc: int) -> None:
    """Exit with rc. On Windows, Qt can crash while its DLLs unload after the
    window has already closed (all saving happens before that), which would
    show a "has stopped working" box. So end the process directly there."""
    if sys.platform == "win32":
        try:
            logging.shutdown()
            sys.stdout and sys.stdout.flush()
            sys.stderr and sys.stderr.flush()
            import ctypes
            k32 = ctypes.windll.kernel32
            k32.TerminateProcess(k32.GetCurrentProcess(), int(rc or 0))
        except Exception:
            pass
    sys.exit(rc)


def main(argv=None):
    from portable_excel.gui.app import main as gui_main
    _exit(gui_main(argv))


if __name__ == "__main__":
    main()
