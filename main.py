import sys
import os
from core.excel_unprotector import unprotect_xlsx


def main():
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
    if not os.path.exists(args.output_file):
        print(f"Error: '{args.output_file}' is missing.")
        sys.exit(1)

    try:
        unprotect_xlsx(args.input_file, args.output_file)
    except Exception as e:
        print(f"Unexpected exception: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
