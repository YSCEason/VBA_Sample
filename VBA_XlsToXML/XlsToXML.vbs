if WScript.Arguments.Count < 2 Then
    WScript.Echo "Error! Please specify the source path and the destination. Usage: XlsToXml SourcePath.xls Destination.xml"
    Wscript.Quit
End If
Dim oExcel
Set oExcel = CreateObject("Excel.Application")
Dim oBook
Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))
oBook.SaveAs WScript.Arguments.Item(1), 46

oBook.Close False
oExcel.Quit