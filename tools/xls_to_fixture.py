"""Convert CIQ .xls exports into JSON fixtures that mirror what Apps Script sees
after Google Drive converts the file to a Google Sheet (Range.getValues()).
Dates become {"$date": "YYYY-MM-DD"}; everything else keeps its value."""
import json, sys, os, datetime, xlrd

def convert(path):
    book = xlrd.open_workbook(path, logfile=open(os.devnull, "w"))
    out = {}
    for sh in book.sheets():
        rows = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    d = xlrd.xldate_as_datetime(cell.value, book.datemode)
                    row.append({"$date": d.strftime("%Y-%m-%d")})
                elif cell.ctype == xlrd.XL_CELL_EMPTY:
                    row.append("")
                else:
                    row.append(cell.value)
            rows.append(row)
        out[sh.name] = rows
    return out

if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    os.makedirs(dst, exist_ok=True)
    for f in sorted(os.listdir(src)):
        if f.lower().endswith(".xls"):
            data = {"fileName": f, "sheets": convert(os.path.join(src, f))}
            with open(os.path.join(dst, f[:-4] + ".json"), "w") as fh:
                json.dump(data, fh)
            print("ok", f)
