Dim oExcel
Set oExcel = CreateObject("Excel.Application")
Dim oBook
Set oBook = oExcel.Workbooks.Open("D:\Easan\MyCoding\MyCoding_Git\Python_ExceltoXML\VBA_XlsToXML\D3S Verification Test Case and Report of LVHWIRP.xlsx")

For i = 1 To oBook.Sheets.Count
    oBook.Sheets(i).SaveAs "D:\Easan\MyCoding\MyCoding_Git\Python_ExceltoXML\VBA_XlsToXML\CSV_TEMP\" & oBook.Sheets(i).Name & ".csv", 6
Next

oBook.Close False
oExcel.Quit