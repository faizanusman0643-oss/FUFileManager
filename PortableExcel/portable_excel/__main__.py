"""Entry point: python -m portable_excel [file]"""
import sys


def main(argv=None):
    from portable_excel.gui.app import main as gui_main
    sys.exit(gui_main(argv))


if __name__ == "__main__":
    main()
