Public Sub GetAlphaVantageStockHD()

Application.ScreenUpdating = False

    Dim Stock_Name As String
    Dim avAPIKey As String
    Dim avFunction As String
    Dim avOutputSize As String
    Dim avDataType As String
    Dim avExtension As String

    Dim AWB As Workbook
    Dim NewSheetName As String
    Dim timestamp As String
    
    Dim CurrentDate As Variant
    Dim PrevYearDate As Variant
    
    Dim RowNum As Integer
    Dim ContParse As Boolean
    
    Dim CheckRange As Range
    Dim MinVal, MaxVal, CurrentVal, AboveVal, PercentVal, OneYearVal, GrowthVal, AverageVal, AboveAveVal As Variant
    
    Set AWB = ThisWorkbook
    
' test code
  
    

'If part is for TSE stocks, you might need to adapt it to your Stock exchange
   
    If Range("B1").Value = "XTSE" Then
        avExtension = ".TO"
        Else
        avExtension = ""
    End If
    


'Parameters definition
    Stock_Name = Range("B2")
    avAPIKey = Range("B7").Value
    avFunction = Range("B4")
    avOutputSize = "full"
    avDataType = "csv"
     
    'APIKey : get it at https://www.alphavantage.co/support/#api-key
    'Function : TIME_SERIES_DAILY_ADJUSTED  and more, check alphavantage documentation
    'OutputSize: compact is latest 100 data points; (default) full is up to 20 years of historical data.
    'DataType :  json (default) or csv
   
' more test
    
   
' Download from Alpha Vantage

    Workbooks.Open Filename:="https://www.alphavantage.co/query?" & _
        "function=" & avFunction & _
        "&symbol=" & Stock_Name & avExtension & _
        "&outputsize=" & avOutputSize & _
        "&apikey=" & avAPIKey & _
        "&datatype=" & avDataType

' Change sheet name
   

    timestamp = Format(Now, "MMddyyyy_hhmmss")
    NewSheetName = Stock_Name & "_" & timestamp
    Workbooks("query").Sheets("query").Name = NewSheetName
   

' move sheet
   
    Sheets(NewSheetName).Move After:=AWB.Sheets(1)
    
' AutoFit All Columns on Worksheet

    ThisWorkbook.Worksheets(NewSheetName).Cells.EntireColumn.AutoFit
    
' Change zoom level

    ThisWorkbook.Worksheets(NewSheetName).Select
    ActiveWindow.Zoom = 110
    
' Find one year back
    CurrentDate = Date
    PrevYearDate = DateAdd("yyyy", -1, CurrentDate)

    RowNum = 2
    ContParse = True
    
    Do While ContParse
    
        If IsDate(ThisWorkbook.Worksheets(NewSheetName).Range("A" & RowNum).Value) = False Then
            ContParse = False
            
            If RowNum > 2 Then
                RowNum = RowNum - 1
            End If
    
        ElseIf ThisWorkbook.Worksheets(NewSheetName).Range("A" & RowNum).Value > PrevYearDate Then
            RowNum = RowNum + 1
            
        Else:
            ContParse = False
        End If
        
    Loop
    
    Set CheckRange = Worksheets(NewSheetName).Range("D2:D" & RowNum)
    MinVal = Application.WorksheetFunction.Min(CheckRange)
    Set CheckRange = Worksheets(NewSheetName).Range("C2:C" & RowNum)
    MaxVal = Application.WorksheetFunction.Max(CheckRange)
    CurrentVal = Worksheets(NewSheetName).Range("E2").Value
    RangeVal = MaxVal - MinVal
    AboveVal = CurrentVal - MinVal
      
    If RangeVal <> 0 Then
        PercentVal = AboveVal / RangeVal
    Else:
        PercentVal = 0
    End If
    
    OneYearVal = Worksheets(NewSheetName).Range("E" & RowNum).Value
    
    If OneYearVal <> 0 Then
        GrowthVal = (CurrentVal - OneYearVal) / OneYearVal
    Else:
        GrowthVal = 0
    End If
    
    Set CheckRange = Worksheets(NewSheetName).Range("E2:E" & RowNum)
    AverageVal = Application.WorksheetFunction.Average(CheckRange)
    
    If AverageVal <> 0 Then
        AboveAveVal = (CurrentVal - AverageVal) / AverageVal
    Else:
        AboveAveVal = 0
    End If
        
    Worksheets(NewSheetName).Range("K1").Value = "Max"
    Worksheets(NewSheetName).Range("K2").Value = MaxVal
    
    Worksheets(NewSheetName).Range("L1").Value = "Min"
    Worksheets(NewSheetName).Range("L2").Value = MinVal
    
    Worksheets(NewSheetName).Range("M1").Value = "Current"
    Worksheets(NewSheetName).Range("M2").Value = CurrentVal
    
    Worksheets(NewSheetName).Range("K4").Value = "Range"
    Worksheets(NewSheetName).Range("K5").Value = RangeVal
    
    Worksheets(NewSheetName).Range("M4").Value = "Above"
    Worksheets(NewSheetName).Range("M5").Value = AboveVal
    
    Worksheets(NewSheetName).Range("K7").Value = "Percent"
    Worksheets(NewSheetName).Range("K8").Value = FormatPercent(PercentVal, 2, vbTrue)

    Worksheets(NewSheetName).Range("N1").Value = "OneYear"
    Worksheets(NewSheetName).Range("N2").Value = OneYearVal
    
    Worksheets(NewSheetName).Range("M7").Value = "Growth"
    Worksheets(NewSheetName).Range("M8").Value = FormatPercent(GrowthVal, 2, vbTrue)
    
    Worksheets(NewSheetName).Range("O4").Value = "Average"
    Worksheets(NewSheetName).Range("O5").Value = AverageVal
    
    Worksheets(NewSheetName).Range("O7").Value = "AboveAve"
    Worksheets(NewSheetName).Range("O8").Value = FormatPercent(AboveAveVal, 2, vbTrue)
       
End Sub

