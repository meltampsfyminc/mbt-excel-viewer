import sys
import json
import argparse
import pandas as pd

def list_sheets(path):
    try:
        xl = pd.ExcelFile(path)
        sheets = []
        for name in xl.sheet_names:
            df = xl.parse(name, nrows=1)
            cols = [str(c) for c in df.columns.tolist()] if not df.empty else []
            sheets.append({"name": name, "columns": cols})
        print(json.dumps(sheets))
        return 0
    except Exception as e:
        print(json.dumps([{"error": str(e)}]))
        return 1

def read_page(path, sheet, page, size):
    try:
        skiprows = (page - 1) * size
        # Read only needed rows to reduce memory usage
        df = pd.read_excel(path, sheet_name=sheet, header=0)
        rows = df.iloc[skiprows: skiprows + size].values.tolist()
        print(json.dumps(rows))
        return 0
    except Exception as e:
        print(json.dumps([["Error", str(e)]]))
        return 1

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--list-sheets', action='store_true')
    parser.add_argument('--read', action='store_true')
    parser.add_argument('path', nargs='?')
    parser.add_argument('--sheet', default=None)
    parser.add_argument('--page', type=int, default=1)
    parser.add_argument('--size', type=int, default=50)
    args = parser.parse_args()

    if args.list_sheets and args.path:
        sys.exit(list_sheets(args.path))
    elif args.read and args.path and args.sheet:
        sys.exit(read_page(args.path, args.sheet, args.page, args.size))
    else:
        print(json.dumps([["Error", "Invalid arguments"]]))
        sys.exit(1)

if __name__ == "__main__":
    main()
