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
    Dim OneYearAdj, GrowthAdj As Variant
    
    Set AWB = ThisWorkbook
    

'If part is for TSE stocks, you might need to adapt it to your Stock exchange
   
    If Range("B1").Value = "XTSE" Then
        avExtension = ".TO"
        Else
        avExtension = ""
    End If
    

'Parameters definition
    Stock_Name = UCase(Range("B2"))
    avAPIKey = Range("B7").Value
    avFunction = Range("B4")
    avOutputSize = "full"
    avDataType = "csv"
     
    'APIKey : get it at https://www.alphavantage.co/support/#api-key
    'Function : TIME_SERIES_DAILY_ADJUSTED  and more, check alphavantage documentation
    'OutputSize: compact is latest 100 data points; (default) full is up to 20 years of historical data.
    'DataType :  json (default) or csv
   

' Download from Alpha Vantage

    Workbooks.Open Filename:="https://www.alphavantage.co/query?" & _
        "function=" & avFunction & _
        "&symbol=" & Stock_Name & avExtension & _
        "&outputsize=" & avOutputSize & _
        "&apikey=" & avAPIKey & _
        "&datatype=" & avDataType

' Change sheet name
   
    'timestamp = Format(Now, "MMddyyyy_hhmmss")
    'NewSheetName = Stock_Name & "_" & timestamp
    
    NewSheetName = Stock_Name & avExtension
    Workbooks("query").Sheets("query").Name = NewSheetName
    
' check if sheet exists and delete

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = NewSheetName Then
            
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            
            Exit For
        End If
    Next
   
' move sheet
   
    Sheets(NewSheetName).Move After:=AWB.Sheets(1)

    
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
    
    ' collect adjusted one year value as well
    OneYearAdj = Worksheets(NewSheetName).Range("F" & RowNum).Value
    
    If OneYearAdj <> 0 Then
        GrowthAdj = (CurrentVal - OneYearAdj) / OneYearAdj
    Else:
        GrowthAdj = 0
    End If
    
    Set CheckRange = Worksheets(NewSheetName).Range("E2:E" & RowNum)
    AverageVal = Application.WorksheetFunction.Average(CheckRange)
    AverageVal = Application.WorksheetFunction.Round(AverageVal, 2)
    
    If AverageVal <> 0 Then
        AboveAveVal = (CurrentVal - AverageVal) / AverageVal
    Else:
        AboveAveVal = 0
    End If
        
    Worksheets(NewSheetName).Range("K1").Value = "MaxVal"
    Worksheets(NewSheetName).Range("K2").Value = MaxVal
    
    Worksheets(NewSheetName).Range("L1").Value = "MinVal"
    Worksheets(NewSheetName).Range("L2").Value = MinVal
    
    Worksheets(NewSheetName).Range("M1").Value = "Current"
    Worksheets(NewSheetName).Range("M2").Value = CurrentVal
    
    Worksheets(NewSheetName).Range("K4").Value = "Range"
    Worksheets(NewSheetName).Range("K5").Value = RangeVal
    
    Worksheets(NewSheetName).Range("M4").Value = "Above"
    Worksheets(NewSheetName).Range("M5").Value = AboveVal
    
    Worksheets(NewSheetName).Range("K7").Value = "MinMax"
    Worksheets(NewSheetName).Range("K8").Value = FormatPercent(PercentVal, 2, vbTrue)

    Worksheets(NewSheetName).Range("N1").Value = "OneYear"
    Worksheets(NewSheetName).Range("N2").Value = OneYearVal
    
    Worksheets(NewSheetName).Range("M7").Value = "1yrGrowth"
    Worksheets(NewSheetName).Range("M8").Value = FormatPercent(GrowthVal, 2, vbTrue)
    
    Worksheets(NewSheetName).Range("O4").Value = "Average"
    Worksheets(NewSheetName).Range("O5").Value = AverageVal
    
    Worksheets(NewSheetName).Range("O7").Value = "AboveAve"
    Worksheets(NewSheetName).Range("O8").Value = FormatPercent(AboveAveVal, 2, vbTrue)
    
    ' Parse entire set of data
    RowNum = 2
    ContParse = True
    
        Do While ContParse
    
        If IsDate(ThisWorkbook.Worksheets(NewSheetName).Range("A" & RowNum).Value) = False Then
            ContParse = False
            
            If RowNum > 2 Then
                RowNum = RowNum - 1
            End If
            
        Else:
            RowNum = RowNum + 1
        End If
        
    Loop
    
    ' print the final row
    Worksheets(NewSheetName).Range("K11").Value = "Max Row"
    Worksheets(NewSheetName).Range("K12").Value = RowNum
    
    ' find and print the first closing price
    Dim firstVal As Variant
    firstVal = ThisWorkbook.Worksheets(NewSheetName).Range("E" & RowNum).Value
    
    Worksheets(NewSheetName).Range("M11").Value = "1st val"
    Worksheets(NewSheetName).Range("M12").Value = firstVal
    
    ' determine total growth over the period
    Dim totGrow As Variant
    totGrow = (CurrentVal - firstVal) / firstVal
    
    Worksheets(NewSheetName).Range("O11").Value = "Tot Grow"
    Worksheets(NewSheetName).Range("O12").Value = FormatPercent(totGrow, 2, vbTrue)
    
    ' find the first date
    Dim firstDate As Date
    firstDate = ThisWorkbook.Worksheets(NewSheetName).Range("A" & RowNum).Value
    
    Worksheets(NewSheetName).Range("K14").Value = "1st date"
    Worksheets(NewSheetName).Range("K15").Value = firstDate
    
    ' find current date
    Dim currDate As Date
    currDate = ThisWorkbook.Worksheets(NewSheetName).Range("A2").Value
    
    Worksheets(NewSheetName).Range("M14").Value = "currDate"
    Worksheets(NewSheetName).Range("M15").Value = currDate
    
    ' find diff dates
    Dim diffDates As Variant
    diffDates = currDate - firstDate
    
    Worksheets(NewSheetName).Range("O14").Value = "diffDates"
    Worksheets(NewSheetName).Range("O15").Value = diffDates
    
    ' total years of data
    Dim totYears As Variant
    totYears = diffDates / 365.2425
    
    Worksheets(NewSheetName).Range("K17").Value = "TotYears"
    Worksheets(NewSheetName).Range("K18").Value = Application.WorksheetFunction.Round(totYears, 2)
    
    ' find the average annual growth
    Dim aveGrow As Variant
    aveGrow = totGrow / totYears
    
    Worksheets(NewSheetName).Range("M17").Value = "AveGrowth"
    Worksheets(NewSheetName).Range("M18").Value = FormatPercent(aveGrow, 2, vbTrue)
    
    ' find the difference between previous year growth and ave 1 year growth
    Dim diffGrow As Variant
    diffGrow = GrowthVal - aveGrow
    
    Worksheets(NewSheetName).Range("O17").Value = "DiffGrowth"
    Worksheets(NewSheetName).Range("O18").Value = FormatPercent(diffGrow, 2, vbTrue)
    
    ' 1-year adjusted price
    Worksheets(NewSheetName).Range("K20").Value = "OneYearAdj"
    Worksheets(NewSheetName).Range("K21").Value = Application.WorksheetFunction.Round(OneYearAdj, 2)
    
    ' 1-year adjusted growth
    Worksheets(NewSheetName).Range("M20").Value = "1yrAdj"
    Worksheets(NewSheetName).Range("M21").Value = FormatPercent(GrowthAdj, 2, vbTrue)
    
    ' find the difference between adjusted growth and price growth
    Dim diffAdj As Variant
    diffAdj = GrowthAdj - GrowthVal
    
    Worksheets(NewSheetName).Range("O20").Value = "DiffAdj"
    Worksheets(NewSheetName).Range("O21").Value = FormatPercent(diffAdj, 2, vbTrue)
        
    ' Formatting
    ' AutoFit All Columns on Worksheet

    ThisWorkbook.Worksheets(NewSheetName).Cells.EntireColumn.AutoFit
    
    ' Change zoom level

    ThisWorkbook.Worksheets(NewSheetName).Select
    ActiveWindow.Zoom = 110
       
End Sub

