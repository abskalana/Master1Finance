Attribute VB_Name = "Repository"
Sub ExtractData(Symbols As String, startDate As String, endDate As String, OutputData As Worksheet)
    Dim resultFromYahoo As String
    Dim objRequest
    Dim csv_rows() As String
    Dim dateArray As Variant
    Dim openArray As Variant
    Dim iRows As Integer
    Dim CSV_Fields As Variant
    Dim tickerURL As String

    tickerURL = "https://query1.finance.yahoo.com/v7/finance/download/" & Symbols & _
        "?period1=" & startDate & _
        "&period2=" & endDate & _
        "&interval=1d & events=history"
               
    Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objRequest
        .Open "GET", tickerURL, False
        .send
        .waitForResponse
        resultFromYahoo = .ResponseText
    End With
    
    csv_rows() = Split(resultFromYahoo, Chr(10))
    csv_rows = Filter(csv_rows, csv_rows(0), False)
    

    ReDim dateArray(0 To UBound(csv_rows), 0 To 0) As Variant
    ReDim openArray(0 To UBound(csv_rows), 0 To 0) As Variant
     
    For iRows = LBound(csv_rows) To UBound(csv_rows)
        CSV_Fields = Split(csv_rows(iRows), ",")
        dateArray(iRows, 0) = CDate(CSV_Fields(0))
        openArray(iRows, 0) = Val(CSV_Fields(1))
    Next iRows
 
    OutputData.Cells(1, 1).Value = "Date"
    OutputData.Cells(1, i + 1).Value = Symbols
    OutputData.Range(OutputData.Cells(2, 1), OutputData.Cells(UBound(dateArray, 1) + 2, 1)).Value = dateArray
    OutputData.Range(OutputData.Cells(2, i + 1), OutputData.Cells(UBound(openArray, 1) + 2, i + 1)).Value = openArray
    
End Sub








