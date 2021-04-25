Sub Main()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Call import_LSE_data
Call import_price
ThisWorkbook.RefreshAll

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub import_price()

Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Price Data (Comp)")
Dim wbTemp As Workbook
Dim wsTemp As Worksheet
Dim lRowMain, lRowTemp As Integer
Dim priceRoot As String: priceRoot = ThisWorkbook.Path & "\stock_pricing_info\"
Dim filePath As String: filePath = Dir(priceRoot & "*.csv")


Application.ScreenUpdating = False

lRowMain = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
ws.Range("A2:J" & lRowMain).Clear

While filePath <> ""
    Set wbTemp = Workbooks.Open(priceRoot & filePath)
    Set wsTemp = wbTemp.Worksheets(Left(filePath, InStr(filePath, ".") - 1))
        lRowTemp = wsTemp.Cells(Rows.Count, "A").End(xlUp).Row
        lRowMain = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    With ws
        .Range(.Cells(lRowMain, 1), .Cells(lRowMain + lRowTemp - 2, 1)).Value = Left(filePath, InStr(filePath, ".") - 1)
        .Range(.Cells(lRowMain, 2), .Cells(lRowMain + lRowTemp - 2, 2)).Value = wsTemp.Range("A2:A" & lRowTemp).Value
        .Range(.Cells(lRowMain, 3), .Cells(lRowMain + lRowTemp - 2, 3)).Value = wsTemp.Range("E2:E" & lRowTemp).Value
        For i = lRowMain + 1 To lRowMain + lRowTemp - 2
            .Cells(i, 4).Value = (.Cells(i, 3).Value / .Cells(lRowMain, 3).Value) - 1
        Next i
    End With
    
    wbTemp.Close
    Debug.Print filePath
    filePath = Dir
Wend

    lRowMain = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    With ws
    .Cells(2, 5).Value = "=VLOOKUP(A2,'Market Cap & Sector'!A:D,3,FALSE)"
    .Cells(2, 6).Value = "=VLOOKUP(A2,'Market Cap & Sector'!A:D,2,FALSE)"
    .Cells(2, 7).Value = "=SUMIFS(F:F,B:B,$B2,E:E,E2)"
    .Cells(2, 8).Value = "=D2*(F2/G2)"
    .Cells(2, 9).Value = "=SUMIF(B:B,$B2,F:F)"
    .Cells(2, 10).Value = "=D2*(F2/I2)"
    ws.Range("A2:A" & lRowMain).Replace what:="-", replacement:=""
    .Range("E2:J2").AutoFill Destination:=.Range("E2:J" & lRowMain)
   End With

End Sub

Sub import_LSE_data()

Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Market Cap & Sector")
Dim wbTemp As Workbook
Dim wsTemp As Worksheet
Dim lseRoot As String: lseRoot = ThisWorkbook.Path & "\"
Dim filePath As String: filePath = Dir(lseRoot)
Dim flag1, flag2 As Boolean
Dim timeout As Integer


lRowMain = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
ws.Range("A2:C" & lRowMain).Clear

Do Until flag1 = True And flag2 = True Or timeout = 10
    If filePath = "sector_info.csv" Then
        Set wbTemp = Workbooks.Open(lseRoot & filePath)
        Set wsTemp = wbTemp.Worksheets(Left(filePath, InStr(filePath, ".") - 1))
            lRowTemp = wsTemp.Cells(Rows.Count, "A").End(xlUp).Row
            
                wsTemp.Range("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
                TrailingMinusNumbers:=True
        
            ws.Range("C2:D" & lRowTemp).Value = wsTemp.Range("B2:C" & lRowTemp).Value
            wbTemp.Close savechanges:=False
        flag1 = True
    End If
    
    If filePath = "stock_info.csv" Then
        Set wbTemp = Workbooks.Open(lseRoot & filePath)
        Set wsTemp = wbTemp.Worksheets(Left(filePath, InStr(filePath, ".") - 1))
            lRowTemp = wsTemp.Cells(Rows.Count, "A").End(xlUp).Row
        
                wsTemp.Range("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1)), _
                TrailingMinusNumbers:=True
        
            ws.Range("A2:B" & lRowTemp).Value = wsTemp.Range("A2:B" & lRowTemp).Value
            wbTemp.Close savechanges:=False
        flag2 = True
    End If
    filePath = Dir
    timout = timeout + 1
Loop


lRowMain = ws.Cells(Rows.Count, "A").End(xlUp).Row
ws.Range("A2:A" & lRowMain).Replace what:=".", replacement:=""

End Sub

