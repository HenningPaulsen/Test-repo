# Import pandas
import pandas as pd
import json
import sys

# Assign spreadsheet filename to `file`
def convertToJson(filename):
    # Load spreadsheet
    filenameExcel = filename + ".xlsx"
    filenameJson = filename + ".json"

    try:
        xls = pd.ExcelFile(filenameExcel)
    except:
        print("Filnavn ikke validt - skal indeholde filnavn uden extension")
        sys.exit()

    print("Converting " + filenameExcel + " to Json")

    # Load the sheet into a DataFrame
    xls_df = xls.parse(xls.sheet_names[0])

    json_data = xls_df.to_json(orient="records")
    parsed = json.loads(json_data)
    y = json.dumps(parsed, indent=4)
    f = open(filenameJson, "w")
    f.write(y)
    f.close()

if len(sys.argv) == 2:
    filename = sys.argv[1]
    convertToJson(filename)
else:
    print("Der skal angives filnavn (uden extension)")