import sys
import os
import io
from core.excel_unprotector import unprotect_xlsx


def main():
    # A redirected stdout on Windows uses the ANSI code page, which lacks some
    # characters, such as Cyrillic in a VBA line: print() would fail after the
    # workbook is written. Escape them with backslashreplace, as stderr does,
    # keeping the encoding. Only a TextIOWrapper has reconfigure: tests replace
    # stdout with a StringIO.
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(errors="backslashreplace")

    import argparse

    parser = argparse.ArgumentParser(
        description="Remove XML and VBA protections in Excel files (.xlsx / xlsm)."
    )

    parser.add_argument("input_file", help="Path to Excel file to unprotect")
    parser.add_argument("output_file", help="Path to Excel file unprotected")

    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        print(f"Error: '{args.input_file}' is missing.")
        sys.exit(1)

    try:
        protect_calls = unprotect_xlsx(args.input_file, args.output_file)
    except Exception as e:
        print(f"Unexpected exception: {e}")
        sys.exit(1)

    if protect_calls:
        print(
            "Warning: VBA macros call Protect and may protect the workbook "
            "again when they run. Remove these lines (the password is often "
            "written in them):"
        )
        for call in protect_calls:
            print(f"  {call.module}:{call.line_number}: {call.line}")


if __name__ == "__main__":
    main()
