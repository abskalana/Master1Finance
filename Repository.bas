Attribute VB_Name = "Repository"
Sub GetData(startDate As String, endDate As String, period As String, Symbols As Variant, OutputData As Worksheet)
    Dim crumb As String
    Dim cookie As String
    Dim validCookieCrumb As Boolean
    
    
    Call getCookieCrumb(crumb, cookie, validCookieCrumb)

    Dim i As Integer
    
    For i = LBound(Symbols) To UBound(Symbols)
        Dim ticker As String
        ticker = Symbols(i)
        Call ExtractData(ticker, startDate, endDate, period, cookie, crumb, i, OutputData)
    Next i
    
End Sub



Sub getCookieCrumb(crumb As String, cookie As String, validCookieCrumb As Boolean)

    Dim i As Integer
    Dim str As String
    Dim crumbStartPos As Long
    Dim crumbEndPos As Long
    Dim objRequest
 
    validCookieCrumb = False
    
    For i = 0 To 5
        Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        With objRequest
            .Open "GET", "https://finance.yahoo.com/lookup?s=bananas", False
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
            .send
            .waitForResponse (30)
            cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
            crumbStartPos = InStrRev(.ResponseText, """crumb"":""") + 9
            crumbEndPos = crumbStartPos + 11
            crumb = Mid(.ResponseText, crumbStartPos, crumbEndPos - crumbStartPos)
        End With
        
        If Len(crumb) = 11 Then
            validCookieCrumb = True
            Exit For
        End If:
        
    Next i
End Sub

Sub ExtractData(Symbols As String, startDate As String, endDate As String, period As String, cookie As String, crumb As String, i As Integer, OutputData As Worksheet)
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
        "&interval=" & period & "&events=history" & "&crumb=" & crumb
               
    Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objRequest
        .Open "GET", tickerURL, False
        .setRequestHeader "Cookie", cookie
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








