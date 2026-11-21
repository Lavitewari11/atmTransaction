import openpyxl

wb = openpyxl.load_workbook("docs/04_Traceability_Matrix.xlsx")
ws = wb["Traceability_Matrix"]

tc_mappings = {
    4:  "TC-F-01, TC-F-04, TC-T-06", # BR-01
    5:  "TC-F-02, TC-F-07, TC-T-04", # BR-02
    6:  "TC-F-10, TC-T-07",          # BR-03
    7:  "TC-F-08, TC-T-05",          # BR-04
    8:  "TC-F-09, TC-T-06",          # BR-05
    9:  "TC-F-13, TC-T-05",          # BR-06
    10: "TC-F-14, TC-T-05",          # BR-07
    11: "TC-F-06, TC-T-09",          # BR-08
    12: "TC-F-12, TC-T-02",          # BR-09
    13: "TC-F-11, TC-T-08",          # BR-10
    14: "TC-F-05, TC-F-16",          # BR-11
    15: "TC-F-15, TC-T-05"           # BR-12
}

for row_idx, tc_str in tc_mappings.items():
    ws.cell(row=row_idx, column=12, value=tc_str)
    # Update Status to Completed
    ws.cell(row=row_idx, column=7, value="Tested")

wb.save("docs/04_Traceability_Matrix.xlsx")
print("RTM updated with Test Cases and status set to Tested.")
