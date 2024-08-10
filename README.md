# Automated_IT_Ticket_Analyzation
This was a macro I created for my supervisor so that he could pull any amount of data from our ticketing system and have 24 different graphs populate. 

**Step1:** The first step was for him to press the first button in the worksheet labeled "control panel". This would create the raw data worksheet that he would then cut the raw data from our ticketing system and paste it into the raw data worksheet.
![whole_view](https://github.com/jbradstr/Automated_IT_Ticket_Analyzation/blob/main/AITA_pic1.png?raw=true)

**Step2:** For the second step he simply had to press the second button which takes about 45 seconds or so, but it takes the 10 columns of the raw data pull and transfroms them into 160 columns of data that is then used to create the different charts. (Note names have been deleted for confidentiality.) See below:
This picture shows the raw data:
![whole_view](https://github.com/jbradstr/Automated_IT_Ticket_Analyzation/blob/main/AITA_pic2_rawdata.png?raw=true)

This next picture is the expanded form of the raw data, but a zoomed out version in excel:
![whole_view](https://github.com/jbradstr/Automated_IT_Ticket_Analyzation/blob/main/AITA_pic6_160rawdata.png?raw=true)

The next picture shows the first charts that are created. These charts show the top 3 highschools, middle schools, and elementary schools based on IT ticket volumne:
![whole_view](https://github.com/jbradstr/Automated_IT_Ticket_Analyzation/blob/main/AITA_pic3_top3_closedtickets.png?raw=true)

The next charts are of two different groups.  My supervisor wanted me to split the data between the IT department and the technical assistants (TAs) to see how many tickets those closed versus us.
![whole_view](https://github.com/jbradstr/Automated_IT_Ticket_Analyzation/blob/main/AITA_pic4_avgclosetime_IT_TA.png?raw=true)

Finally I was able to create charts that showed the top 5 tickets by type and then by year for each of the top 3 high, middle and elementary schools. I incorporated chatGPT to help me create the charts so that the colors for each ticket type corresponded to every other ticket type throughout each of the graphs.  For instance, "CB-Other" is shaded dark blue and I was able to keep this uniform throughout each of the charts.
![whole_view](https://github.com/jbradstr/Automated_IT_Ticket_Analyzation/blob/main/AITA_pic5_bytypebyyear.png?raw=true)

Feel free to view the code here:
```vbscript

Public Sub create_raw_data_ws()
    
    Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "raw_data"
    ActiveSheet.Move After:=Worksheets(Worksheets.count)
    
    cp.Activate
    
End Sub

Public Sub step2()

    Call optimize
    Call delete_comma_A
    Call delete_comma_B
    Call insert_date_diff
    Call turn_on_MSRR
    Call avg_tk_close_time_for_data_pulled
    Call avg_tk_close_time_by_year_ALL
    Call rawdata_FN_LN_concat
    Call TA_IT_split_data
    Call avg_tk_close_time_by_year_TAs
    Call avg_tk_close_time_by_year_ITd
    Call TA_IT_percent_tk_closed_by_year
    
    Call closed_tks_by_school_by_year
    Call count_yearly_tks_types_CRH
    Call count_yearly_tks_types_FMH
    Call count_yearly_tks_types_NFH
    Call count_yearly_tks_types_FMM
    Call count_yearly_tks_types_GHM
    Call count_yearly_tks_types_PKM
    Call count_yearly_tks_types_KTE
    Call count_yearly_tks_types_RVE
    Call count_yearly_tks_types_FME
    
    Call create_IT_Dashboard
    
    Call yearly_avg_close_time_chart1
    Call TA_yearly_avg_close_time_chart2
    Call TA_yearly_tsks_closed_chart3
    Call IT_yearly_avg_close_time_chart4
    Call IT_yearly_tsks_closed_chart5
    Call IT_TA_percent_closed_chart6
    
    Call school_charts
    Call schools_by_tsk_type
    Call ColorBarsByDescriptionAcrossCharts
    
End Sub

Sub turn_on_MSRR()

    'enable microsoft scripting runtime reference
    Dim ref As Object
    Dim sLibName As String
    sLibName = "Scripting"
    
    ' Check if the reference is already set
    For Each ref In ThisWorkbook.VBProject.References
        If ref.Name = sLibName Then
            'MsgBox sLibName & " reference is already set."
            Exit Sub
        End If
    Next ref
    
    ' Add the reference
    ThisWorkbook.VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
    
    'MsgBox sLibName & " reference has been added."

End Sub

Public Sub delete_comma_A()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim targetColumn As String
    rd_lastRow = rd.Range("A" & rd.Rows.count).End(xlUp).Row
    Dim i As Long
    
    ' Set the worksheet to the active sheet
    'Set rd = ActiveSheet
    
    ' Define the target column (change this to the column you need)
    targetColumn = "A"

    ' Loop through each cell in the target column and remove commas
    For i = 2 To rd_lastRow
    
        If InStr(rd.Cells(i, targetColumn).Value, ",") > 0 Then
            rd.Cells(i, targetColumn).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
            rd.Cells(i, targetColumn).Value = Replace(rd.Cells(i, targetColumn).Value, ",", "")
        End If
    
    Next i
    
    rd.Activate
    rd.Columns.AutoFit

End Sub


Public Sub delete_comma_B()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim targetColumn As String
    rd_lastRow = rd.Range("B" & rd.Rows.count).End(xlUp).Row
    Dim i As Long
    
    ' Set the worksheet to the active sheet
    'Set rd = ActiveSheet
    
    ' Define the target column (change this to the column you need)
    targetColumn = "B"

    
    ' Loop through each cell in the target column and remove commas
    For i = 2 To rd_lastRow
        If InStr(rd.Cells(i, targetColumn).Value, ",") > 0 Then
            rd.Cells(i, targetColumn).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
            rd.Cells(i, targetColumn).Value = Replace(rd.Cells(i, targetColumn).Value, ",", "")
        End If
    Next i
    
    rd.Columns.AutoFit
    
End Sub


Public Sub insert_date_diff()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("A" & rd.Rows.count).End(xlUp).Row
    Dim i As Long
    
    ' Set the worksheet to the active sheet
    'set rd = ActiveSheet
    
    ' Insert new column and format it for number
    rd.Columns("C").Insert Shift:=xlToRight
    rd.Columns(3).NumberFormat = "0.0000"
    
    
    ' Set headers for the new column
    rd.Cells(1, 3).Value = "Date Difference"
    
    
    ' Loop through each row and calculate the difference
    For i = 2 To rd_lastRow
        
        If IsDate(rd.Cells(i, 1).Value) And IsDate(rd.Cells(i, 2).Value) Then
            Dim startDate As Date
            Dim endDate As Date
            Dim diff As Double
            Dim days As Long
            Dim minutes As Long
            
            startDate = rd.Cells(i, 2).Value
            endDate = rd.Cells(i, 1).Value
            diff = endDate - startDate
            
            ' Output the days and minutes as numbers in separate columns
            
            rd.Cells(i, 3).Value = diff
            
        Else
        
            rd.Cells(i, 3).Value = "Invalid Date"
        
        End If
    
    Next i

End Sub

Public Sub avg_tk_close_time_for_data_pulled()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("A" & rd.Rows.count).End(xlUp).Row
    
    'avg ticket close time all data
    rd.Cells(1, 13) = "Average Ticket Close Time For Pulled Data"
    rd.Cells(2, 13) = WorksheetFunction.Average(Range("C2:C" & rd_lastRow))

End Sub

Public Sub avg_tk_close_time_by_year_ALL()

    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("A" & rd.Rows.count).End(xlUp).Row
    
    rd.Range("A1").CurrentRegion.Sort _
        key1:=rd.Range("A1"), Order1:=xlAscending, Header:=xlYes
    
    Dim yearCol As Long
    Dim dataCol As Long
    Dim outputYearCol As Long
    Dim outputAvgCol As Long
    Dim i As Long
    Dim yearDict As Object
    Dim currentYear As Variant
    Dim yearSum As Double
    Dim yearCount As Long
    Dim outputRow As Long
    
    'set col numbers
    yearCol = 1
    dataCol = 3
    outputYearCol = 15
    outputAvgTskCloseCol = 16
    
    'initialize dictionary
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    'loop through each row to sum data and count occurrences for each year
    For i = 2 To rd_lastRow
        ' Extract the year based on the timeframe from August 1st to July 31st of the following year
        If Month(rd.Cells(i, yearCol).Value) < 8 Then
            currentYear = Year(rd.Cells(i, yearCol).Value) - 1
        Else
            currentYear = Year(rd.Cells(i, yearCol).Value)
        End If
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = Array(0, 0) ' Initialize sum and count
        End If
        yearSum = yearDict(currentYear)(0) + rd.Cells(i, dataCol).Value
        yearCount = yearDict(currentYear)(1) + 1
        yearDict(currentYear) = Array(yearSum, yearCount)
    Next i
        
    ' Output the averages for each year
    outputRow = 2
    For Each currentYear In yearDict.Keys
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        If yearDict(currentYear)(1) <> 0 Then
            rd.Cells(outputRow, outputAvgTskCloseCol).Value = yearDict(currentYear)(0) / yearDict(currentYear)(1)
        Else
            rd.Cells(outputRow, outputAvgTskCloseCol).Value = "N/A" ' Handle division by zero
        End If
        outputRow = outputRow + 1
    Next currentYear
    
    ' AutoFit the new columns
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputAvgTskCloseCol).AutoFit
    
    rd.Cells(1, 15) = "School Year"
    rd.Cells(1, 16) = "Avg Ticket Close Time"
    
End Sub


Public Sub rawdata_FN_LN_concat()

    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim rd_lastRow As Long
    Dim firstNameCol As Long
    Dim lastNameCol As Long
    Dim fullNameCol As Long
    Dim i As Long
    rd_lastRow = rd.Range("A" & rd.Rows.count).End(xlUp).Row
    
    rd.Columns(10).Insert
    firstNameCol = 8
    lastNameCol = 9
    fullNameCol = 10
    
    rd.Cells(1, 10) = "Full Name"
    
    For i = 2 To rd_lastRow ' Assuming the data starts from the second row
        rd.Cells(i, fullNameCol).Value = rd.Cells(i, firstNameCol).Value & " " & rd.Cells(i, lastNameCol).Value
    Next i
    
    rd.Columns(10).AutoFit
    'concatenate first and last name to separate column

End Sub


Public Sub TA_IT_FN_LN_concat()

    Dim tait As Worksheet: Set tait = ThisWorkbook.Worksheets("TAs_ITd_Sites")
    Dim rd_lastRow As Long
    Dim firstNameCol As Long
    Dim lastNameCol As Long
    Dim fullNameCol As Long
    Dim i As Long
    rd_lastRow = tait.Range("A" & tait.Rows.count).End(xlUp).Row
    
    'rd.Columns(10).Insert
    firstNameCol = 4
    lastNameCol = 5
    fullNameCol = 6
    
    'rd.Cells(1, 10) = "Full Name"
    
    For i = 2 To rd_lastRow ' Assuming the data starts from the second row
        tait.Cells(i, fullNameCol).Value = tait.Cells(i, firstNameCol).Value & " " & tait.Cells(i, lastNameCol).Value
    Next i
    
    tait.Columns(fullNameCol).AutoFit
    'concatenate first and last name to separate column

    
End Sub


Public Sub TA_IT_split_data()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim tait As Worksheet: Set tait = ThisWorkbook.Worksheets("TAs_ITd_Sites")
    rd_lastRow = rd.Range("A" & rd.Rows.count).End(xlUp).Row
    Dim destrow As Long
    Dim destrow1 As Long
    Dim destcol As Long
    Dim destcol1 As Long
    
    rd.Activate
    destrow = 2
    destcol = 31
    destrow1 = 2
    destcol1 = 19
    
    rd.Range("A1:K1").Copy
    rd.Cells(1, 19).PasteSpecial
    rd.Cells(1, 31).PasteSpecial
    rd.Columns(19).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Columns(20).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Columns(31).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Columns(32).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    
    
    For i = 2 To rd_lastRow
        
        'TA name list
        full_name = Application.Match(rd.Cells(i, 10), tait.Range(tait.Cells(2, 3), tait.Cells(1000, 3)), 0)
    
        If IsError(full_name) Then
            
            'IT name list
            full_name = Application.Match(rd.Cells(i, 10), tait.Range(tait.Cells(2, 6), tait.Cells(1000, 6)), 0)
            
            If IsError(full_name) Then
                
                resumeloop = MsgBox("Last name " & rd.Cells(i, 10).Value & " not in TA or IT lists." & vbNewLine _
                & "Do you want to continue to the next iteration?", vbYesNo, "Error") = vbYes
                If Not resumeloop Then
                    Exit Sub
                End If
            
            ElseIf full_name > 0 Then
                
                For j = 1 To 11
                
                    rd.Cells(destrow, destcol) = rd.Cells(i, j)
                    destcol = destcol + 1
                
                Next j
                
                destcol = 31
                destrow = destrow + 1
            
            End If
                
        ElseIf full_name > 0 Then
            
            For k = 1 To 11
                
                rd.Cells(destrow1, destcol1) = rd.Cells(i, k)
                destcol1 = destcol1 + 1
                
            Next k
            
            destcol1 = 19
            destrow1 = destrow1 + 1
    
        End If
    
    Next i
    
    rd.Columns(19).AutoFit
    rd.Columns(20).AutoFit
    rd.Columns(31).AutoFit
    rd.Columns(32).AutoFit
    
    rd.Columns(30).Insert
    rd.Columns(30).Insert
    rd.Columns(30).Insert
    
    
    
End Sub


'
'Public Sub remove_dups()
'
'    Dim tait As Worksheet: Set tait = ThisWorkbook.Worksheets("Sheet4")
'    tait.Range("A:B").RemoveDuplicates Columns:=1, Header:=xlYes
'
'End Sub


Public Sub avg_tk_close_time_by_year_TAs()

    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("S" & rd.Rows.count).End(xlUp).Row
    
    rd.Range("S1").CurrentRegion.Sort _
        key1:=rd.Range("S1"), Order1:=xlAscending, Header:=xlYes
    
    Dim yearCol As Long
    Dim dataCol As Long
    Dim outputYearCol As Long
    Dim outputAvgCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearDict As Object
    Dim currentYear As Variant
    Dim yearSum As Double
    Dim yearCount As Long
    Dim outputRow As Long
    
    'set col numbers
    yearCol = 19
    dataCol = 21
    outputYearCol = 30
    outputAvgTskCloseCol = 31
    outputCountCol = 32
    
    'initialize dictionary
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    'loop through each row to sum data and count occurrences for each year
    For i = 2 To rd_lastRow
        ' Extract the year based on the timeframe from August 1st to July 31st of the following year
        If Month(rd.Cells(i, yearCol).Value) < 8 Then
            currentYear = Year(rd.Cells(i, yearCol).Value) - 1
        Else
            currentYear = Year(rd.Cells(i, yearCol).Value)
        End If
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = Array(0, 0) ' Initialize sum and count
        End If
        yearSum = yearDict(currentYear)(0) + rd.Cells(i, dataCol).Value
        yearCount = yearDict(currentYear)(1) + 1
        yearDict(currentYear) = Array(yearSum, yearCount)
    Next i
        
    ' Output the averages for each year
    outputRow = 2
    For Each currentYear In yearDict.Keys
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        If yearDict(currentYear)(1) <> 0 Then
            rd.Cells(outputRow, outputAvgTskCloseCol).Value = yearDict(currentYear)(0) / yearDict(currentYear)(1)
            rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)(1)
        Else
            rd.Cells(outputRow, outputAvgTskCloseCol).Value = "N/A" ' Handle division by zero
            rd.Cells(outputRow, outputCountCol).Value = 0
        End If
        outputRow = outputRow + 1
    Next currentYear
    
    rd.Cells(1, 30) = "School Year"
    rd.Cells(1, 31) = "Avg Ticket Close Time"
    rd.Cells(1, 32) = "Total Tickets Closed"
    
    ' AutoFit the new columns
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputAvgTskCloseCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
End Sub


Public Sub avg_tk_close_time_by_year_ITd()

    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("AH" & rd.Rows.count).End(xlUp).Row
    
    rd.Range("AH1").CurrentRegion.Sort _
        key1:=rd.Range("AH1"), Order1:=xlAscending, Header:=xlYes
    
    Dim yearCol As Long
    Dim dataCol As Long
    Dim outputYearCol As Long
    Dim outputAvgCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearDict As Object
    Dim currentYear As Variant
    Dim yearSum As Double
    Dim yearCount As Long
    Dim outputRow As Long
    
    'set col numbers
    yearCol = 34
    dataCol = 36
    outputYearCol = 45
    outputAvgTskCloseCol = 46
    outputCountCol = 47
    
    'initialize dictionary
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    'loop through each row to sum data and count occurrences for each year
    For i = 2 To rd_lastRow
        ' Extract the year based on the timeframe from August 1st to July 31st of the following year
        If Month(rd.Cells(i, yearCol).Value) < 8 Then
            currentYear = Year(rd.Cells(i, yearCol).Value) - 1
        Else
            currentYear = Year(rd.Cells(i, yearCol).Value)
        End If
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = Array(0, 0) ' Initialize sum and count
        End If
        yearSum = yearDict(currentYear)(0) + rd.Cells(i, dataCol).Value
        yearCount = yearDict(currentYear)(1) + 1
        yearDict(currentYear) = Array(yearSum, yearCount)
    Next i
        
    ' Output the averages for each year
    outputRow = 2
    For Each currentYear In yearDict.Keys
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        If yearDict(currentYear)(1) <> 0 Then
            rd.Cells(outputRow, outputAvgTskCloseCol).Value = yearDict(currentYear)(0) / yearDict(currentYear)(1)
            rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)(1)
        Else
            rd.Cells(outputRow, outputAvgTskCloseCol).Value = "N/A" ' Handle division by zero
            rd.Cells(outputRow, outputCountCol).Value = 0
        End If
        outputRow = outputRow + 1
    Next currentYear
    
    rd.Cells(1, 45) = "School Year"
    rd.Cells(1, 46) = "Avg Ticket Close Time"
    rd.Cells(1, 47) = "Total Tickets Closed"
    
    ' AutoFit the new columns
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputAvgTskCloseCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
End Sub

Public Sub TA_IT_percent_tk_closed_by_year()

    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("AF" & rd.Rows.count).End(xlUp).Row
    
    Dim ta_tsks As Long
    Dim it_tsks As Long
    Dim tsks_total As Long
    Dim percent_TA As Double
    Dim percent_IT As Double
    
    rd.Columns(33).Insert
    rd.Cells(1, 33) = "TA Percent Closed"
    rd.Cells(1, 49) = "IT Percent Closed"
    
    ta_tsks = 32
    it_tsks = 48
    
    For i = 2 To rd_lastRow
    
        tsks_total = rd.Cells(i, ta_tsks) + rd.Cells(i, it_tsks)
        percent_TA = (rd.Cells(i, ta_tsks) / tsks_total)
        percent_IT = (rd.Cells(i, it_tsks) / tsks_total)
    
        rd.Cells(i, 33) = percent_TA
        rd.Cells(i, 49) = percent_IT
    
    Next i

    rd.Columns(33).NumberFormat = "0.00"
    rd.Columns(49).NumberFormat = "0.00"

End Sub


Public Sub optimize()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.EnableAnimations = False
    Application.DisplayStatusBar = False
    Application.PrintCommunication = False

End Sub

'Closed Tickets by School YoY
Public Sub closed_tks_by_school_by_year()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim tait As Worksheet: Set tait = ThisWorkbook.Worksheets("TAs_ITd_Sites")
    rd_lastRow = rd.Range("A" & rd.Rows.count).End(xlUp).Row
    Dim destrow As Long
    Dim destcol As Long
    
    Dim yearCol As Long
    Dim outputYearCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearDict As Object
    Dim currentYear As Variant
    Dim yearCount As Long
    Dim outputRow As Long
    
    'CRHS_____________________________________________________________________________
    rd.Activate
    destrow = 2
    destcol = 51
    
    yearCol = 51
    outputYearCol = 54
    outputCountCol = 55
    
    rd.Columns(51).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Cells(1, 1).Copy
    rd.Cells(1, 51).PasteSpecial
    rd.Cells(1, 5).Copy
    rd.Cells(1, 52).PasteSpecial
    rd.Cells(1, 6).Copy
    rd.Cells(1, 53).PasteSpecial
    
    
    site_name = "CRHS"
    
    rd.Activate
    
    For i = 2 To rd_lastRow
        
        If rd.Cells(i, 5) = site_name Then
            
            rd.Cells(destrow, destcol) = rd.Cells(i, 1)
            rd.Cells(destrow, destcol + 1) = rd.Cells(i, 5)
            rd.Cells(destrow, destcol + 2) = rd.Cells(i, 6)
            
            destcol = 51
            destrow = destrow + 1
        
        Else
        End If
        
   Next i
    
    
    rd.Columns(51).NumberFormat = "YYYY"
    rd.Columns(51).AutoFit
    
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol).Value)
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = 0
        End If
        
        yearCount = yearDict(currentYear) + 1
        yearDict(currentYear) = yearCount
    
    Next i
    
    outputRow = 2
    For Each currentYear In yearDict.Keys
        
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)
        outputRow = outputRow + 1
        
    Next currentYear
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
    
    'FMHS_____________________________________________________________________________
    rd.Activate
    destrow = 2
    destcol = 57
    
    yearCol = 57
    outputYearCol = 60
    outputCountCol = 61
    
    rd.Columns(destcol).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Cells(1, 1).Copy
    rd.Cells(1, destcol).PasteSpecial
    rd.Cells(1, 5).Copy
    rd.Cells(1, destcol + 1).PasteSpecial
    rd.Cells(1, 6).Copy
    rd.Cells(1, destcol + 2).PasteSpecial
    
    
    site_name = "FMHS"
    
    rd.Activate
    
    For i = 2 To rd_lastRow
        
        If rd.Cells(i, 5) = site_name Then
            
            rd.Cells(destrow, destcol) = rd.Cells(i, 1)
            rd.Cells(destrow, destcol + 1) = rd.Cells(i, 5)
            rd.Cells(destrow, destcol + 2) = rd.Cells(i, 6)
            
            destcol = 57
            destrow = destrow + 1
        
        Else
        End If
        
   Next i
    
    
    rd.Columns(destcol).NumberFormat = "YYYY"
    rd.Columns(destcol).AutoFit
    
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol).Value)
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = 0
        End If
        
        yearCount = yearDict(currentYear) + 1
        yearDict(currentYear) = yearCount
    
    Next i
    
    outputRow = 2
    For Each currentYear In yearDict.Keys
        
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)
        outputRow = outputRow + 1
        
    Next currentYear
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
    
    'NFHS_____________________________________________________________________________
    rd.Activate
    destrow = 2
    destcol = 63
    
    yearCol = 63
    outputYearCol = 66
    outputCountCol = 67
    
    rd.Columns(destcol).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Cells(1, 1).Copy
    rd.Cells(1, destcol).PasteSpecial
    rd.Cells(1, 5).Copy
    rd.Cells(1, destcol + 1).PasteSpecial
    rd.Cells(1, 6).Copy
    rd.Cells(1, destcol + 2).PasteSpecial
    
    
    site_name = "NFHS"
    
    rd.Activate
    
    For i = 2 To rd_lastRow
        
        If rd.Cells(i, 5) = site_name Then
            
            rd.Cells(destrow, destcol) = rd.Cells(i, 1)
            rd.Cells(destrow, destcol + 1) = rd.Cells(i, 5)
            rd.Cells(destrow, destcol + 2) = rd.Cells(i, 6)
            
            destcol = 63
            destrow = destrow + 1
        
        Else
        End If
        
   Next i
    
    
    rd.Columns(destcol).NumberFormat = "YYYY"
    rd.Columns(destcol).AutoFit
    
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol).Value)
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = 0
        End If
        
        yearCount = yearDict(currentYear) + 1
        yearDict(currentYear) = yearCount
    
    Next i
    
    outputRow = 2
    For Each currentYear In yearDict.Keys
        
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)
        outputRow = outputRow + 1
        
    Next currentYear
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
    
    'FMMS_____________________________________________________________________________
    rd.Activate
    destrow = 2
    destcol = 69
    
    yearCol = 69
    outputYearCol = 72
    outputCountCol = 73
    
    rd.Columns(destcol).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Cells(1, 1).Copy
    rd.Cells(1, destcol).PasteSpecial
    rd.Cells(1, 5).Copy
    rd.Cells(1, destcol + 1).PasteSpecial
    rd.Cells(1, 6).Copy
    rd.Cells(1, destcol + 2).PasteSpecial
    
    
    site_name = "FMMS"
    
    rd.Activate
    
    For i = 2 To rd_lastRow
        
        If rd.Cells(i, 5) = site_name Then
            
            rd.Cells(destrow, destcol) = rd.Cells(i, 1)
            rd.Cells(destrow, destcol + 1) = rd.Cells(i, 5)
            rd.Cells(destrow, destcol + 2) = rd.Cells(i, 6)
            
            destcol = 69
            destrow = destrow + 1
        
        Else
        End If
        
   Next i
    
    
    rd.Columns(destcol).NumberFormat = "YYYY"
    rd.Columns(destcol).AutoFit
    
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol).Value)
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = 0
        End If
        
        yearCount = yearDict(currentYear) + 1
        yearDict(currentYear) = yearCount
    
    Next i
    
    outputRow = 2
    For Each currentYear In yearDict.Keys
        
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)
        outputRow = outputRow + 1
        
    Next currentYear
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
    
    'GHMS_____________________________________________________________________________
    rd.Activate
    destrow = 2
    destcol = 75
    
    yearCol = 75
    outputYearCol = 78
    outputCountCol = 79
    
    rd.Columns(destcol).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Cells(1, 1).Copy
    rd.Cells(1, destcol).PasteSpecial
    rd.Cells(1, 5).Copy
    rd.Cells(1, destcol + 1).PasteSpecial
    rd.Cells(1, 6).Copy
    rd.Cells(1, destcol + 2).PasteSpecial
    
    
    site_name = "GHMS"
    
    rd.Activate
    
    For i = 2 To rd_lastRow
        
        If rd.Cells(i, 5) = site_name Then
            
            rd.Cells(destrow, destcol) = rd.Cells(i, 1)
            rd.Cells(destrow, destcol + 1) = rd.Cells(i, 5)
            rd.Cells(destrow, destcol + 2) = rd.Cells(i, 6)
            
            destcol = 75
            destrow = destrow + 1
        
        Else
        End If
        
   Next i
    
    
    rd.Columns(destcol).NumberFormat = "YYYY"
    rd.Columns(destcol).AutoFit
    
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol).Value)
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = 0
        End If
        
        yearCount = yearDict(currentYear) + 1
        yearDict(currentYear) = yearCount
    
    Next i
    
    outputRow = 2
    For Each currentYear In yearDict.Keys
        
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)
        outputRow = outputRow + 1
        
    Next currentYear
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputCountCol).AutoFit

    
    'PKMS_____________________________________________________________________________
    rd.Activate
    destrow = 2
    destcol = 81
    
    yearCol = 81
    outputYearCol = 84
    outputCountCol = 85
    
    rd.Columns(destcol).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Cells(1, 1).Copy
    rd.Cells(1, destcol).PasteSpecial
    rd.Cells(1, 5).Copy
    rd.Cells(1, destcol + 1).PasteSpecial
    rd.Cells(1, 6).Copy
    rd.Cells(1, destcol + 2).PasteSpecial
    
    
    site_name = "PKMS"
    
    rd.Activate
    
    For i = 2 To rd_lastRow
        
        If rd.Cells(i, 5) = site_name Then
            
            rd.Cells(destrow, destcol) = rd.Cells(i, 1)
            rd.Cells(destrow, destcol + 1) = rd.Cells(i, 5)
            rd.Cells(destrow, destcol + 2) = rd.Cells(i, 6)
            
            destcol = 81
            destrow = destrow + 1
        
        Else
        End If
        
   Next i
    
    
    rd.Columns(destcol).NumberFormat = "YYYY"
    rd.Columns(destcol).AutoFit
    
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol).Value)
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = 0
        End If
        
        yearCount = yearDict(currentYear) + 1
        yearDict(currentYear) = yearCount
    
    Next i
    
    outputRow = 2
    For Each currentYear In yearDict.Keys
        
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)
        outputRow = outputRow + 1
        
    Next currentYear
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
    
    'KTES_____________________________________________________________________________
    rd.Activate
    destrow = 2
    destcol = 87
    
    yearCol = 87
    outputYearCol = 90
    outputCountCol = 91
    
    rd.Columns(destcol).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Cells(1, 1).Copy
    rd.Cells(1, destcol).PasteSpecial
    rd.Cells(1, 5).Copy
    rd.Cells(1, destcol + 1).PasteSpecial
    rd.Cells(1, 6).Copy
    rd.Cells(1, destcol + 2).PasteSpecial
    
    
    site_name = "KTES"
    
    rd.Activate
    
    For i = 2 To rd_lastRow
        
        If rd.Cells(i, 5) = site_name Then
            
            rd.Cells(destrow, destcol) = rd.Cells(i, 1)
            rd.Cells(destrow, destcol + 1) = rd.Cells(i, 5)
            rd.Cells(destrow, destcol + 2) = rd.Cells(i, 6)
            
            destcol = 87
            destrow = destrow + 1
        
        Else
        End If
        
   Next i
    
    
    rd.Columns(destcol).NumberFormat = "YYYY"
    rd.Columns(destcol).AutoFit
    
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol).Value)
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = 0
        End If
        
        yearCount = yearDict(currentYear) + 1
        yearDict(currentYear) = yearCount
    
    Next i
    
    outputRow = 2
    For Each currentYear In yearDict.Keys
        
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)
        outputRow = outputRow + 1
        
    Next currentYear
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
    
    'RVES_____________________________________________________________________________
    rd.Activate
    destrow = 2
    destcol = 93
    
    yearCol = 93
    outputYearCol = 96
    outputCountCol = 97
    
    rd.Columns(destcol).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Cells(1, 1).Copy
    rd.Cells(1, destcol).PasteSpecial
    rd.Cells(1, 5).Copy
    rd.Cells(1, destcol + 1).PasteSpecial
    rd.Cells(1, 6).Copy
    rd.Cells(1, destcol + 2).PasteSpecial
    
    
    site_name = "RVES"
    
    rd.Activate
    
    For i = 2 To rd_lastRow
        
        If rd.Cells(i, 5) = site_name Then
            
            rd.Cells(destrow, destcol) = rd.Cells(i, 1)
            rd.Cells(destrow, destcol + 1) = rd.Cells(i, 5)
            rd.Cells(destrow, destcol + 2) = rd.Cells(i, 6)
            
            destcol = 93
            destrow = destrow + 1
        
        Else
        End If
        
   Next i
    
    
    rd.Columns(destcol).NumberFormat = "YYYY"
    rd.Columns(destcol).AutoFit
    
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol).Value)
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = 0
        End If
        
        yearCount = yearDict(currentYear) + 1
        yearDict(currentYear) = yearCount
    
    Next i
    
    outputRow = 2
    For Each currentYear In yearDict.Keys
        
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)
        outputRow = outputRow + 1
        
    Next currentYear
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
    
    'FMES_____________________________________________________________________________
    rd.Activate
    destrow = 2
    destcol = 99
    
    yearCol = 99
    outputYearCol = 102
    outputCountCol = 103
    
    rd.Columns(destcol).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    rd.Cells(1, 1).Copy
    rd.Cells(1, destcol).PasteSpecial
    rd.Cells(1, 5).Copy
    rd.Cells(1, destcol + 1).PasteSpecial
    rd.Cells(1, 6).Copy
    rd.Cells(1, destcol + 2).PasteSpecial
    
    
    site_name = "FMES"
    
    rd.Activate
    
    For i = 2 To rd_lastRow
        
        If rd.Cells(i, 5) = site_name Then
            
            rd.Cells(destrow, destcol) = rd.Cells(i, 1)
            rd.Cells(destrow, destcol + 1) = rd.Cells(i, 5)
            rd.Cells(destrow, destcol + 2) = rd.Cells(i, 6)
            
            destcol = 99
            destrow = destrow + 1
        
        Else
        End If
        
   Next i
    
    
    rd.Columns(destcol).NumberFormat = "YYYY"
    rd.Columns(destcol).AutoFit
    
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol).Value)
        
        If Not yearDict.Exists(currentYear) Then
            yearDict(currentYear) = 0
        End If
        
        yearCount = yearDict(currentYear) + 1
        yearDict(currentYear) = yearCount
    
    Next i
    
    outputRow = 2
    For Each currentYear In yearDict.Keys
        
        rd.Cells(outputRow, outputYearCol).Value = currentYear
        rd.Cells(outputRow, outputCountCol).Value = yearDict(currentYear)
        outputRow = outputRow + 1
        
    Next currentYear
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
    
    rd.Columns(56).Insert
    rd.Columns(56).Insert
    rd.Columns(56).Insert
    rd.Columns(65).Insert
    rd.Columns(65).Insert
    rd.Columns(65).Insert
    rd.Columns(74).Insert
    rd.Columns(74).Insert
    rd.Columns(74).Insert
    rd.Columns(83).Insert
    rd.Columns(83).Insert
    rd.Columns(83).Insert
    rd.Columns(92).Insert
    rd.Columns(92).Insert
    rd.Columns(92).Insert
    rd.Columns(101).Insert
    rd.Columns(101).Insert
    rd.Columns(101).Insert
    rd.Columns(110).Insert
    rd.Columns(110).Insert
    rd.Columns(110).Insert
    rd.Columns(119).Insert
    rd.Columns(119).Insert
    rd.Columns(119).Insert

End Sub

Public Sub count_yearly_tks_types_CRH()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("BA" & rd.Rows.count).End(xlUp).Row
    Dim yearCol As Long
    Dim ticketTypeCol As Long
    Dim outputYearCol As Long
    Dim outputTicketTypeCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearTicketDict As Object
    Dim currentYear As Variant
    Dim currentTicketType As String
    Dim ticketCount As Long
    Dim outputRow As Long
    Dim key As Variant
    
    yearCol = 51
    ticketTypeCol = 53
    outputYearCol = 56
    outputTicketTypeCol = 57
    outputCountCol = 58
    'rd.Columns(56).NumberFormat = "General"
    
    Set yearTicketDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol))
            
         ' Ensure the value is treated as a date and extract the year
        currentTicketType = rd.Cells(i, ticketTypeCol).Value
        
        ' Create a unique key for the year-ticket type pair
        key = CStr(currentYear) & "|" & currentTicketType
        
        If Not yearTicketDict.Exists(key) Then
            yearTicketDict(key) = 0 'initialize count
        End If
        
        ticketCount = yearTicketDict(key) + 1
        yearTicketDict(key) = ticketCount
    
    Next i
    
    outputRow = 2
    
    For Each key In yearTicketDict.Keys
    
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        rd.Cells(outputRow, outputYearCol).Value = CInt(keyParts(0))
        rd.Cells(outputRow, outputTicketTypeCol).Value = keyParts(1)
        rd.Cells(outputRow, outputCountCol).Value = yearTicketDict(key)
        
        outputRow = outputRow + 1
        
    Next key
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputTicketTypeCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
        

End Sub

Public Sub count_yearly_tks_types_FMH()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("BH" & rd.Rows.count).End(xlUp).Row
    Dim yearCol As Long
    Dim ticketTypeCol As Long
    Dim outputYearCol As Long
    Dim outputTicketTypeCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearTicketDict As Object
    Dim currentYear As Variant
    Dim currentTicketType As String
    Dim ticketCount As Long
    Dim outputRow As Long
    Dim key As Variant
    
    yearCol = 60
    ticketTypeCol = 62
    outputYearCol = 65
    outputTicketTypeCol = 66
    outputCountCol = 67
    'rd.Columns(56).NumberFormat = "General"
    
    Set yearTicketDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol))
            
         ' Ensure the value is treated as a date and extract the year
        currentTicketType = rd.Cells(i, ticketTypeCol).Value
        
        ' Create a unique key for the year-ticket type pair
        key = CStr(currentYear) & "|" & currentTicketType
        
        If Not yearTicketDict.Exists(key) Then
            yearTicketDict(key) = 0 'initialize count
        End If
        
        ticketCount = yearTicketDict(key) + 1
        yearTicketDict(key) = ticketCount
    
    Next i
    
    outputRow = 2
    
    For Each key In yearTicketDict.Keys
    
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        rd.Cells(outputRow, outputYearCol).Value = CInt(keyParts(0))
        rd.Cells(outputRow, outputTicketTypeCol).Value = keyParts(1)
        rd.Cells(outputRow, outputCountCol).Value = yearTicketDict(key)
        
        outputRow = outputRow + 1
        
    Next key
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputTicketTypeCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
        

End Sub

Public Sub count_yearly_tks_types_NFH()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("BQ" & rd.Rows.count).End(xlUp).Row
    Dim yearCol As Long
    Dim ticketTypeCol As Long
    Dim outputYearCol As Long
    Dim outputTicketTypeCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearTicketDict As Object
    Dim currentYear As Variant
    Dim currentTicketType As String
    Dim ticketCount As Long
    Dim outputRow As Long
    Dim key As Variant
    
    yearCol = 69
    ticketTypeCol = 71
    outputYearCol = 74
    outputTicketTypeCol = 75
    outputCountCol = 76
    'rd.Columns(56).NumberFormat = "General"
    
    Set yearTicketDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol))
            
         ' Ensure the value is treated as a date and extract the year
        currentTicketType = rd.Cells(i, ticketTypeCol).Value
        
        ' Create a unique key for the year-ticket type pair
        key = CStr(currentYear) & "|" & currentTicketType
        
        If Not yearTicketDict.Exists(key) Then
            yearTicketDict(key) = 0 'initialize count
        End If
        
        ticketCount = yearTicketDict(key) + 1
        yearTicketDict(key) = ticketCount
    
    Next i
    
    outputRow = 2
    
    For Each key In yearTicketDict.Keys
    
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        rd.Cells(outputRow, outputYearCol).Value = CInt(keyParts(0))
        rd.Cells(outputRow, outputTicketTypeCol).Value = keyParts(1)
        rd.Cells(outputRow, outputCountCol).Value = yearTicketDict(key)
        
        outputRow = outputRow + 1
        
    Next key
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputTicketTypeCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
        

End Sub

Public Sub count_yearly_tks_types_FMM()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("BZ" & rd.Rows.count).End(xlUp).Row
    Dim yearCol As Long
    Dim ticketTypeCol As Long
    Dim outputYearCol As Long
    Dim outputTicketTypeCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearTicketDict As Object
    Dim currentYear As Variant
    Dim currentTicketType As String
    Dim ticketCount As Long
    Dim outputRow As Long
    Dim key As Variant
    
    yearCol = 78
    ticketTypeCol = 80
    outputYearCol = 83
    outputTicketTypeCol = 84
    outputCountCol = 85
    'rd.Columns(56).NumberFormat = "General"
    
    Set yearTicketDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol))
            
         ' Ensure the value is treated as a date and extract the year
        currentTicketType = rd.Cells(i, ticketTypeCol).Value
        
        ' Create a unique key for the year-ticket type pair
        key = CStr(currentYear) & "|" & currentTicketType
        
        If Not yearTicketDict.Exists(key) Then
            yearTicketDict(key) = 0 'initialize count
        End If
        
        ticketCount = yearTicketDict(key) + 1
        yearTicketDict(key) = ticketCount
    
    Next i
    
    outputRow = 2
    
    For Each key In yearTicketDict.Keys
    
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        rd.Cells(outputRow, outputYearCol).Value = CInt(keyParts(0))
        rd.Cells(outputRow, outputTicketTypeCol).Value = keyParts(1)
        rd.Cells(outputRow, outputCountCol).Value = yearTicketDict(key)
        
        outputRow = outputRow + 1
        
    Next key
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputTicketTypeCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
        

End Sub

Public Sub count_yearly_tks_types_GHM()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("CI" & rd.Rows.count).End(xlUp).Row
    Dim yearCol As Long
    Dim ticketTypeCol As Long
    Dim outputYearCol As Long
    Dim outputTicketTypeCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearTicketDict As Object
    Dim currentYear As Variant
    Dim currentTicketType As String
    Dim ticketCount As Long
    Dim outputRow As Long
    Dim key As Variant
    
    yearCol = 87
    ticketTypeCol = 89
    outputYearCol = 92
    outputTicketTypeCol = 93
    outputCountCol = 94
    'rd.Columns(56).NumberFormat = "General"
    
    Set yearTicketDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol))
            
         ' Ensure the value is treated as a date and extract the year
        currentTicketType = rd.Cells(i, ticketTypeCol).Value
        
        ' Create a unique key for the year-ticket type pair
        key = CStr(currentYear) & "|" & currentTicketType
        
        If Not yearTicketDict.Exists(key) Then
            yearTicketDict(key) = 0 'initialize count
        End If
        
        ticketCount = yearTicketDict(key) + 1
        yearTicketDict(key) = ticketCount
    
    Next i
    
    outputRow = 2
    
    For Each key In yearTicketDict.Keys
    
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        rd.Cells(outputRow, outputYearCol).Value = CInt(keyParts(0))
        rd.Cells(outputRow, outputTicketTypeCol).Value = keyParts(1)
        rd.Cells(outputRow, outputCountCol).Value = yearTicketDict(key)
        
        outputRow = outputRow + 1
        
    Next key
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputTicketTypeCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
        

End Sub

Public Sub count_yearly_tks_types_PKM()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("CR" & rd.Rows.count).End(xlUp).Row
    Dim yearCol As Long
    Dim ticketTypeCol As Long
    Dim outputYearCol As Long
    Dim outputTicketTypeCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearTicketDict As Object
    Dim currentYear As Variant
    Dim currentTicketType As String
    Dim ticketCount As Long
    Dim outputRow As Long
    Dim key As Variant
    
    yearCol = 96
    ticketTypeCol = 98
    outputYearCol = 101
    outputTicketTypeCol = 102
    outputCountCol = 103
    'rd.Columns(56).NumberFormat = "General"
    
    Set yearTicketDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol))
            
         ' Ensure the value is treated as a date and extract the year
        currentTicketType = rd.Cells(i, ticketTypeCol).Value
        
        ' Create a unique key for the year-ticket type pair
        key = CStr(currentYear) & "|" & currentTicketType
        
        If Not yearTicketDict.Exists(key) Then
            yearTicketDict(key) = 0 'initialize count
        End If
        
        ticketCount = yearTicketDict(key) + 1
        yearTicketDict(key) = ticketCount
    
    Next i
    
    outputRow = 2
    
    For Each key In yearTicketDict.Keys
    
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        rd.Cells(outputRow, outputYearCol).Value = CInt(keyParts(0))
        rd.Cells(outputRow, outputTicketTypeCol).Value = keyParts(1)
        rd.Cells(outputRow, outputCountCol).Value = yearTicketDict(key)
        
        outputRow = outputRow + 1
        
    Next key
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputTicketTypeCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
        

End Sub

Public Sub count_yearly_tks_types_KTE()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("DA" & rd.Rows.count).End(xlUp).Row
    Dim yearCol As Long
    Dim ticketTypeCol As Long
    Dim outputYearCol As Long
    Dim outputTicketTypeCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearTicketDict As Object
    Dim currentYear As Variant
    Dim currentTicketType As String
    Dim ticketCount As Long
    Dim outputRow As Long
    Dim key As Variant
    
    yearCol = 105
    ticketTypeCol = 107
    outputYearCol = 110
    outputTicketTypeCol = 111
    outputCountCol = 112
    'rd.Columns(56).NumberFormat = "General"
    
    Set yearTicketDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol))
            
         ' Ensure the value is treated as a date and extract the year
        currentTicketType = rd.Cells(i, ticketTypeCol).Value
        
        ' Create a unique key for the year-ticket type pair
        key = CStr(currentYear) & "|" & currentTicketType
        
        If Not yearTicketDict.Exists(key) Then
            yearTicketDict(key) = 0 'initialize count
        End If
        
        ticketCount = yearTicketDict(key) + 1
        yearTicketDict(key) = ticketCount
    
    Next i
    
    outputRow = 2
    
    For Each key In yearTicketDict.Keys
    
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        rd.Cells(outputRow, outputYearCol).Value = CInt(keyParts(0))
        rd.Cells(outputRow, outputTicketTypeCol).Value = keyParts(1)
        rd.Cells(outputRow, outputCountCol).Value = yearTicketDict(key)
        
        outputRow = outputRow + 1
        
    Next key
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputTicketTypeCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
        

End Sub

Public Sub count_yearly_tks_types_RVE()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("DJ" & rd.Rows.count).End(xlUp).Row
    Dim yearCol As Long
    Dim ticketTypeCol As Long
    Dim outputYearCol As Long
    Dim outputTicketTypeCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearTicketDict As Object
    Dim currentYear As Variant
    Dim currentTicketType As String
    Dim ticketCount As Long
    Dim outputRow As Long
    Dim key As Variant
    
    yearCol = 114
    ticketTypeCol = 116
    outputYearCol = 119
    outputTicketTypeCol = 120
    outputCountCol = 121
    'rd.Columns(56).NumberFormat = "General"
    
    Set yearTicketDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol))
            
         ' Ensure the value is treated as a date and extract the year
        currentTicketType = rd.Cells(i, ticketTypeCol).Value
        
        ' Create a unique key for the year-ticket type pair
        key = CStr(currentYear) & "|" & currentTicketType
        
        If Not yearTicketDict.Exists(key) Then
            yearTicketDict(key) = 0 'initialize count
        End If
        
        ticketCount = yearTicketDict(key) + 1
        yearTicketDict(key) = ticketCount
    
    Next i
    
    outputRow = 2
    
    For Each key In yearTicketDict.Keys
    
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        rd.Cells(outputRow, outputYearCol).Value = CInt(keyParts(0))
        rd.Cells(outputRow, outputTicketTypeCol).Value = keyParts(1)
        rd.Cells(outputRow, outputCountCol).Value = yearTicketDict(key)
        
        outputRow = outputRow + 1
        
    Next key
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputTicketTypeCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
        

End Sub

Public Sub count_yearly_tks_types_FME()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd_lastRow = rd.Range("DS" & rd.Rows.count).End(xlUp).Row
    Dim yearCol As Long
    Dim ticketTypeCol As Long
    Dim outputYearCol As Long
    Dim outputTicketTypeCol As Long
    Dim outputCountCol As Long
    Dim i As Long
    Dim yearTicketDict As Object
    Dim currentYear As Variant
    Dim currentTicketType As String
    Dim ticketCount As Long
    Dim outputRow As Long
    Dim key As Variant
    
    yearCol = 123
    ticketTypeCol = 125
    outputYearCol = 128
    outputTicketTypeCol = 129
    outputCountCol = 130
    'rd.Columns(56).NumberFormat = "General"
    
    Set yearTicketDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To rd_lastRow
        
        currentYear = Year(rd.Cells(i, yearCol))
            
         ' Ensure the value is treated as a date and extract the year
        currentTicketType = rd.Cells(i, ticketTypeCol).Value
        
        ' Create a unique key for the year-ticket type pair
        key = CStr(currentYear) & "|" & currentTicketType
        
        If Not yearTicketDict.Exists(key) Then
            yearTicketDict(key) = 0 'initialize count
        End If
        
        ticketCount = yearTicketDict(key) + 1
        yearTicketDict(key) = ticketCount
    
    Next i
    
    outputRow = 2
    
    For Each key In yearTicketDict.Keys
    
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        rd.Cells(outputRow, outputYearCol).Value = CInt(keyParts(0))
        rd.Cells(outputRow, outputTicketTypeCol).Value = keyParts(1)
        rd.Cells(outputRow, outputCountCol).Value = yearTicketDict(key)
        
        outputRow = outputRow + 1
        
    Next key
    
    rd.Columns(outputYearCol).AutoFit
    rd.Columns(outputTicketTypeCol).AutoFit
    rd.Columns(outputCountCol).AutoFit
        

End Sub


Public Sub create_IT_Dashboard()
    
    Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "IT_Dashboard"
    ActiveSheet.Move After:=Worksheets(Worksheets.count)
    
    Dim itd As Worksheet: Set itd = ThisWorkbook.Worksheets("IT_Dashboard")
    
    itd.Activate
    
End Sub


Public Sub chart1()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim itd As Worksheet: Set itd = ThisWorkbook.Worksheets("IT_Dashboard")
    Dim rng As Range
    Dim cht1 As ChartObject
    Dim lastRow As Long
    
    'lastrow = rd.Cells(rd.Rows.Count, "Q").End(xlUp).Row
    Set rng = rd.Range("P1:Q6" & lastRow)

    
    itd.Activate
    
    Set cht1 = itd.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    
    With cht1.Chart
        .SetSourceData Source:=rng
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Avg Ticket Close Time"
        .ChartStyle = 222
        .HasLegend = False
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Avg Time"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Year"
        .ApplyDataLabels
    End With
    
    With cht1
        .Left = .Left - 51.75
        .Top = .Top - 157.5
        .Width = .Width * 1.0645833333
        .Height = .Height * 1.5243055556
    End With


End Sub

Public Sub test_date_format()
    
    'tool macro
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    
    'rd.Columns(56).Numberformat =
    rd.Cells(2, 56) = Year(rd.Cells(2, 51))


End Sub

Public Sub delete_cols()
    
    'tool macro
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    rd.Columns("AY:CQM").Delete
    
End Sub

Public Sub unhide_statusbar()
    
    'tool macro
    Application.DisplayStatusBar = True
    
End Sub


Public Sub yearly_avg_close_time_chart1()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim itd As Worksheet: Set itd = ThisWorkbook.Worksheets("IT_Dashboard")
    rd_lastRowP = rd.Range("P" & rd.Rows.count).End(xlUp).Row
    rd_lastRowQ = rd.Range("Q" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("P2:P" & rd_lastRowP)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = itd.ChartObjects.Add(Left:=25, Width:=375, Top:=0, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Avg Close Time"
        .SeriesCollection(1).XValues = rd.Range("P2:P" & rd_lastRowP)
        .SeriesCollection(1).values = rd.Range("Q2:Q" & rd_lastRowQ)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0.00"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "Yearly Avg Close Time"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Avg Close Time"
    End With
    
End Sub


Public Sub TA_yearly_avg_close_time_chart2()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim itd As Worksheet: Set itd = ThisWorkbook.Worksheets("IT_Dashboard")
    rd_lastRowAD = rd.Range("AD" & rd.Rows.count).End(xlUp).Row
    rd_lastRowAE = rd.Range("AE" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("AD2:AD" & rd_lastRowAD)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = itd.ChartObjects.Add(Left:=25, Width:=375, Top:=235, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Avg Close Time"
        .SeriesCollection(1).XValues = rd.Range("AD2:AD" & rd_lastRowAD)
        .SeriesCollection(1).values = rd.Range("AE2:AE" & rd_lastRowAE)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0.00"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "TA Yearly Avg Close Time"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Avg Close Time (days)"
    End With
    
End Sub


Public Sub TA_yearly_tsks_closed_chart3()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim itd As Worksheet: Set itd = ThisWorkbook.Worksheets("IT_Dashboard")
    rd_lastRowAD = rd.Range("AD" & rd.Rows.count).End(xlUp).Row
    rd_lastRowAF = rd.Range("AF" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("AD2:AD" & rd_lastRowAD)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = itd.ChartObjects.Add(Left:=25, Width:=375, Top:=470, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tickets Closed"
        .SeriesCollection(1).XValues = rd.Range("AD2:AD" & rd_lastRowAD)
        .SeriesCollection(1).values = rd.Range("AF2:AF" & rd_lastRowAF)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "TA Yearly Tickets Closed"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
    End With
    
End Sub

Public Sub IT_yearly_avg_close_time_chart4()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim itd As Worksheet: Set itd = ThisWorkbook.Worksheets("IT_Dashboard")
    rd_lastRowAD = rd.Range("AD" & rd.Rows.count).End(xlUp).Row
    rd_lastRowAU = rd.Range("AU" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("AD2:AD" & rd_lastRowAD)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = itd.ChartObjects.Add(Left:=410, Width:=375, Top:=235, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Avg Close Time"
        .SeriesCollection(1).XValues = rd.Range("AD2:AD" & rd_lastRowAD)
        .SeriesCollection(1).values = rd.Range("AU2:AU" & rd_lastRowAU)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0.00"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "IT Yearly Avg Close Time"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Avg Close Time (days)"
    End With
    
End Sub


Public Sub IT_yearly_tsks_closed_chart5()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim itd As Worksheet: Set itd = ThisWorkbook.Worksheets("IT_Dashboard")
    rd_lastRowAD = rd.Range("AD" & rd.Rows.count).End(xlUp).Row
    rd_lastRowAV = rd.Range("AV" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("AD2:AD" & rd_lastRowAD)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = itd.ChartObjects.Add(Left:=410, Width:=375, Top:=470, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tickets Closed"
        .SeriesCollection(1).XValues = rd.Range("AD2:AD" & rd_lastRowAD)
        .SeriesCollection(1).values = rd.Range("AV2:AV" & rd_lastRowAV)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "IT Yearly Tickets Closed"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
    End With
    
End Sub


Sub IT_TA_percent_closed_chart6()
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim itd As Worksheet: Set itd = ThisWorkbook.Worksheets("IT_Dashboard")
    rd_lastRowYear = rd.Range("AT" & rd.Rows.count).End(xlUp).Row
    rd_lastRowTA = rd.Range("AG" & rd.Rows.count).End(xlUp).Row
    rd_lastRowIT = rd.Range("AW" & rd.Rows.count).End(xlUp).Row
    
    Dim chartObj As ChartObject
    
    Set chartObj = itd.ChartObjects.Add(Left:=410, Width:=375, Top:=0, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered
    
        ' Add first series (col AG)
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "TA"
        .SeriesCollection(1).XValues = rd.Range("AT2:AT" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("AG2:AG" & rd_lastRowTA)
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).dataLabels.ShowValue = True
        .SeriesCollection(1).dataLabels.NumberFormat = "0%"
        .SeriesCollection(1).dataLabels.Font.Size = 8
        
        ' Add second series (col AW)
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "IT"
        .SeriesCollection(2).XValues = rd.Range("AT2:AT" & rd_lastRowYear)
        .SeriesCollection(2).values = rd.Range("AW2:AW" & rd_lastRowIT)
        .SeriesCollection(2).HasDataLabels = True
        .SeriesCollection(2).dataLabels.ShowValue = True
        .SeriesCollection(2).dataLabels.NumberFormat = "0%"
        .SeriesCollection(2).dataLabels.Font.Size = 8
        
        'format y-axis to show percents from 0 - 100%
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .Axes(xlValue).HasMajorGridlines = False ' Remove major gridlines for value axis

        ' Remove gridlines for the category axis (x-axis)
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlCategory).HasMinorGridlines = False
        
        'add chart title
        .HasTitle = True
        .ChartTitle.Text = "Closed Ticket Percentages by Year"
        
        'format series colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 112, 192) ' Blue
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 0, 0) ' Red
        
        'add axis titles
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Percent Closed"

    End With
    
    itd.Activate
        
End Sub


Sub reset_graph()

    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim itd As Worksheet: Set itd = ThisWorkbook.Worksheets("IT_Dashboard")
    Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
    Dim response As VbMsgBoxResult
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    
    ' Prompt user for confirmation
    response = MsgBox("Do you want to start over?", vbYesNo + vbQuestion, "Delete Sheets")
    
    ' If user selects Yes, delete the sheets
    If response = vbYes Then
        Application.DisplayAlerts = False
        rd.Delete
        itd.Delete
        sc.Delete
        std.Delete
        sbtby.Delete
        Application.DisplayAlerts = True
    Else
        MsgBox "Operation cancelled.", vbInformation
    End If
    
End Sub








'#######################################################################################




Public Sub school_charts()

    Call Sort_year_type_count_for_schools
    Call create_school_comp_ws
    Call CRH_tsk_count_chart1
    Call FMH_tsk_count_chart2
    Call NFH_tsk_count_chart3
    Call FMM_tsk_count_chart4
    Call GHM_tsk_count_chart5
    Call PKM_tsk_count_chart6
    Call KTE_tsk_count_chart7
    Call RVE_tsk_count_chart8
    Call FME_tsk_count_chart9

End Sub


Public Sub create_school_comp_ws()

    Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
        
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "School_Comparison"
    ActiveSheet.Move After:=Worksheets(Worksheets.count)
    
    cp.Activate
    
End Sub


Public Sub Sort_year_type_count_for_schools()
    
    'CRH____________________________________________________________________________
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    
    rd_lastRow = rd.Cells(rd.Rows.count, "BD").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With rd.Sort
        .SortFields.Clear
        .SortFields.Add key:=rd.Range("BD2:BD" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=rd.Range("BF2:BF" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange rd.Range("BD1:BF" & rd_lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    
    End With
    
    
    
    'FMH____________________________________________________________________________
    
    
    rd_lastRow = rd.Cells(rd.Rows.count, "BM").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With rd.Sort
        .SortFields.Clear
        .SortFields.Add key:=rd.Range("BM2:BM" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=rd.Range("BO2:BO" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange rd.Range("BM1:BO" & rd_lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    
    End With
    
    
        
    'NFH____________________________________________________________________________
    
    
    rd_lastRow = rd.Cells(rd.Rows.count, "BV").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With rd.Sort
        .SortFields.Clear
        .SortFields.Add key:=rd.Range("BV2:BV" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=rd.Range("BX2:BX" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange rd.Range("BV1:BX" & rd_lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    
    End With
    
    
        
    'FMM____________________________________________________________________________
    
    
    rd_lastRow = rd.Cells(rd.Rows.count, "CE").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With rd.Sort
        .SortFields.Clear
        .SortFields.Add key:=rd.Range("CE2:CE" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=rd.Range("CG2:CG" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange rd.Range("CE1:CG" & rd_lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    
    End With
    
    
    'GHM____________________________________________________________________________
    
    
    rd_lastRow = rd.Cells(rd.Rows.count, "CN").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With rd.Sort
        .SortFields.Clear
        .SortFields.Add key:=rd.Range("CN2:CN" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=rd.Range("CP2:CP" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange rd.Range("CN1:CP" & rd_lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    
    End With
    
    
    'PKM____________________________________________________________________________
    
    
    rd_lastRow = rd.Cells(rd.Rows.count, "CW").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With rd.Sort
        .SortFields.Clear
        .SortFields.Add key:=rd.Range("CW2:CW" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=rd.Range("CY2:CY" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange rd.Range("CW1:CY" & rd_lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    
    End With
    
    
    'KTE____________________________________________________________________________
    
    
    rd_lastRow = rd.Cells(rd.Rows.count, "DF").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With rd.Sort
        .SortFields.Clear
        .SortFields.Add key:=rd.Range("DF2:DF" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=rd.Range("DH2:DH" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange rd.Range("DF1:DH" & rd_lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    
    End With
    
    
    'RVE____________________________________________________________________________
    
    
    rd_lastRow = rd.Cells(rd.Rows.count, "DO").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With rd.Sort
        .SortFields.Clear
        .SortFields.Add key:=rd.Range("DO2:DO" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=rd.Range("DQ2:DQ" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange rd.Range("DO1:DQ" & rd_lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    
    End With
    
    
    'FME____________________________________________________________________________
    
    
    rd_lastRow = rd.Cells(rd.Rows.count, "DX").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With rd.Sort
        .SortFields.Clear
        .SortFields.Add key:=rd.Range("DX2:DX" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=rd.Range("DZ2:DZ" & rd_lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange rd.Range("DX1:DZ" & rd_lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    
    End With
    
End Sub


Public Sub CRH_tsk_count_chart1()
    
    'CRH ----------------------------------------------------------------------------
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    rd_lastRowYear = rd.Range("BB" & rd.Rows.count).End(xlUp).Row
    rd_lastRowCount = rd.Range("BC" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("BB2:BB" & rd_lastRowYear - 1)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = sc.ChartObjects.Add(Left:=25, Width:=375, Top:=0, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tasks Closed"
        .SeriesCollection(1).XValues = rd.Range("BB2:BB" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("BC2:BC" & rd_lastRowCount - 1)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "CRH Closed Tickets by Year"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Closed Tickets"
    
    End With


End Sub

Public Sub FMH_tsk_count_chart2()
    
    'FMH----------------------------------------------------------------------------
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    rd_lastRowYear = rd.Range("BK" & rd.Rows.count).End(xlUp).Row
    rd_lastRowCount = rd.Range("BL" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("BK2:BK" & rd_lastRowYear - 1)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = sc.ChartObjects.Add(Left:=410, Width:=375, Top:=0, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tasks Closed"
        .SeriesCollection(1).XValues = rd.Range("BK2:BK" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("BL2:BL" & rd_lastRowCount - 1)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "FMH Closed Tickets by Year"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Closed Tickets"
    
    End With


End Sub


Public Sub NFH_tsk_count_chart3()
    
    'NFH----------------------------------------------------------------------------
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    rd_lastRowYear = rd.Range("BT" & rd.Rows.count).End(xlUp).Row
    rd_lastRowCount = rd.Range("BU" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("BT2:BT" & rd_lastRowYear - 1)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = sc.ChartObjects.Add(Left:=795, Width:=375, Top:=0, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tasks Closed"
        .SeriesCollection(1).XValues = rd.Range("BT2:BT" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("BU2:BU" & rd_lastRowCount - 1)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "NFH Closed Tickets by Year"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Closed Tickets"
    
    End With


End Sub


Public Sub FMM_tsk_count_chart4()
    
    'FMM----------------------------------------------------------------------------
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    rd_lastRowYear = rd.Range("CC" & rd.Rows.count).End(xlUp).Row
    rd_lastRowCount = rd.Range("CD" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("CC2:CC" & rd_lastRowYear - 1)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = sc.ChartObjects.Add(Left:=25, Width:=375, Top:=235, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tasks Closed"
        .SeriesCollection(1).XValues = rd.Range("CC2:CC" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("CD2:CD" & rd_lastRowCount - 1)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "FMM Closed Tickets by Year"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Closed Tickets"
    
    End With


End Sub

Public Sub GHM_tsk_count_chart5()
    
    'GHM----------------------------------------------------------------------------
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    rd_lastRowYear = rd.Range("CL" & rd.Rows.count).End(xlUp).Row
    rd_lastRowCount = rd.Range("CM" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("CL2:CL" & rd_lastRowYear - 1)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = sc.ChartObjects.Add(Left:=410, Width:=375, Top:=235, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tasks Closed"
        .SeriesCollection(1).XValues = rd.Range("CL2:CL" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("CM2:CM" & rd_lastRowCount - 1)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "GHM Closed Tickets by Year"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Closed Tickets"
    
    End With


End Sub


Public Sub PKM_tsk_count_chart6()
    
    'PKM----------------------------------------------------------------------------
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    rd_lastRowYear = rd.Range("CU" & rd.Rows.count).End(xlUp).Row
    rd_lastRowCount = rd.Range("CV" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("CU2:CU" & rd_lastRowYear - 1)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = sc.ChartObjects.Add(Left:=795, Width:=375, Top:=235, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tasks Closed"
        .SeriesCollection(1).XValues = rd.Range("CU2:CU" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("CV2:CV" & rd_lastRowCount - 1)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "PKM Closed Tickets by Year"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Closed Tickets"
    
    End With


End Sub

Public Sub KTE_tsk_count_chart7()
    
    'KTE----------------------------------------------------------------------------
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    rd_lastRowYear = rd.Range("DD" & rd.Rows.count).End(xlUp).Row
    rd_lastRowCount = rd.Range("DE" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("DD2:DD" & rd_lastRowYear - 1)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = sc.ChartObjects.Add(Left:=25, Width:=375, Top:=470, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tasks Closed"
        .SeriesCollection(1).XValues = rd.Range("DD2:DD" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("DE2:DE" & rd_lastRowCount - 1)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "KTE Closed Tickets by Year"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Closed Tickets"
    
    End With


End Sub


Public Sub RVE_tsk_count_chart8()
    
    'RVE----------------------------------------------------------------------------
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    rd_lastRowYear = rd.Range("DM" & rd.Rows.count).End(xlUp).Row
    rd_lastRowCount = rd.Range("DN" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("DM2:DM" & rd_lastRowYear - 1)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = sc.ChartObjects.Add(Left:=410, Width:=375, Top:=470, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tasks Closed"
        .SeriesCollection(1).XValues = rd.Range("DM2:DM" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("DN2:DN" & rd_lastRowCount - 1)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "RVE Closed Tickets by Year"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Closed Tickets"
    
    End With


End Sub


Public Sub FME_tsk_count_chart9()
    
    'FME----------------------------------------------------------------------------
    
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    Dim sc As Worksheet: Set sc = ThisWorkbook.Worksheets("School_Comparison")
    rd_lastRowYear = rd.Range("DV" & rd.Rows.count).End(xlUp).Row
    rd_lastRowCount = rd.Range("DW" & rd.Rows.count).End(xlUp).Row

    ' Define the data
    Dim years As Variant
    years = rd.Range("DV2:DV" & rd_lastRowYear - 1)
    Dim i As Integer

    ' Create the bar chart
    Dim chartObj As ChartObject
    Set chartObj = sc.ChartObjects.Add(Left:=795, Width:=375, Top:=470, Height:=225)
    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' Remove any existing series
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add the new series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Tasks Closed"
        .SeriesCollection(1).XValues = rd.Range("DV2:DV" & rd_lastRowYear)
        .SeriesCollection(1).values = rd.Range("DW2:DW" & rd_lastRowCount - 1)
        .SeriesCollection(1).HasDataLabels = True
        
        ' Format the data labels to 2 decimal places
        Dim dataLabels As dataLabels
        Set dataLabels = .SeriesCollection(1).dataLabels
        Dim lbl As DataLabel
        For Each lbl In dataLabels
            lbl.NumberFormat = "0"
        Next lbl

        ' Apply Chart Style 209
        .ChartStyle = 209

        ' Remove the legend
        .HasLegend = False

        .HasTitle = True
        .ChartTitle.Text = "FME Closed Tickets by Year"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Years"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Closed Tickets"
    
    End With


End Sub



'#######################################################################################



Public Sub schools_by_tsk_type()

    Call create_school_by_tsk_type
    Call CRH_by_type_data
    Call FMH_by_type_data
    Call NFH_by_type_data
    Call FMM_by_type_data
    Call GHM_by_type_data
    Call PKM_by_type_data
    Call KTE_by_type_data
    Call RVE_by_type_data
    Call FME_by_type_data
    Call sort_year_std
    Call std_delete_cell_contents
    Call CRH_by_year_by_type_chart
    Call FMH_by_year_by_type_chart
    Call NFH_by_year_by_type_chart
    Call FMM_by_year_by_type_chart
    Call PKM_by_year_by_type_chart
    Call GHM_by_year_by_type_chart
    Call KTE_by_year_by_type_chart
    Call RVE_by_year_by_type_chart
    Call FME_by_year_by_type_chart

End Sub

Public Sub color_graphs()

    Call crh_color_bars_by_description
    Call fmh_color_bars_by_description
    Call nfh_color_bars_by_description
    Call fmm_color_bars_by_description
    Call pkm_color_bars_by_description
    Call ghm_color_bars_by_description
    Call kte_color_bars_by_description
    Call rve_color_bars_by_description
    Call fme_color_bars_by_description
    

End Sub


Public Sub create_school_by_tsk_type()

    Dim cp As Worksheet: Set cp = ThisWorkbook.Worksheets("Control_Panel")
     
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "school_type_data"
    ActiveSheet.Move After:=Worksheets(Worksheets.count)
     
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Schools_by_type_by_year"
    ActiveSheet.Move After:=Worksheets(Worksheets.count)
    
    cp.Activate
    
End Sub


Public Sub CRH_by_type_data()
    
    'CRHS______________________________________________________________________________
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    crh_lastrow = rd.Range("bd" & rd.Rows.count).End(xlUp).Row
    Dim count As Integer
    
    count = 0
    destrow = 2
    rd_start_col = 56
    destcol = 1
    
        
    
        
        For i = 2 To crh_lastrow
            
            If rd.Cells(i, rd_start_col) = rd.Cells(i + 1, rd_start_col) And count < 5 Then
                
                For j = 1 To 3
                    
                    std.Cells(destrow, destcol) = rd.Cells(i, 55 + destcol)
                    destcol = destcol + 1
                    
                Next j
                
                destcol = 1
                count = count + 1
                destrow = destrow + 1
                
            ElseIf rd.Cells(i, rd_start_col) <> rd.Cells(i + 1, rd_start_col) Then
                
                If IsEmpty(rd.Cells(i + 1, rd_start_col)) Then Exit For
                
                count = 0
            
            End If
            
        Next i
    
    std.Cells(1, 1) = "CRHS"
    
End Sub

Public Sub FMH_by_type_data()

    'FMH______________________________________________________________________________
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    crh_lastrow = rd.Range("bm" & rd.Rows.count).End(xlUp).Row
    Dim count As Integer
    
    count = 0
    destrow = 2
    rd_start_col = 65
    destcol = 5
        
        For i = 2 To crh_lastrow
            
            If rd.Cells(i, rd_start_col) = rd.Cells(i + 1, rd_start_col) And count < 5 Then
                
                For j = 1 To 3
                    
                    std.Cells(destrow, destcol) = rd.Cells(i, 60 + destcol)
                    destcol = destcol + 1
                    
                Next j
                
                destcol = 5
                count = count + 1
                destrow = destrow + 1
                
            ElseIf rd.Cells(i, rd_start_col) <> rd.Cells(i + 1, rd_start_col) Then
                
                If IsEmpty(rd.Cells(i + 1, rd_start_col)) Then Exit For
                
                count = 0
            
            End If
            
        Next i
    
    std.Cells(1, 5) = "FMHS"

End Sub


Public Sub NFH_by_type_data()

    'NFH______________________________________________________________________________
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    crh_lastrow = rd.Range("bv" & rd.Rows.count).End(xlUp).Row
    Dim count As Integer
    
    count = 0
    destrow = 2
    rd_start_col = 74
    destcol = 9
        
        For i = 2 To crh_lastrow
            
            If rd.Cells(i, rd_start_col) = rd.Cells(i + 1, rd_start_col) And count < 5 Then
                
                For j = 1 To 3
                    
                    std.Cells(destrow, destcol) = rd.Cells(i, 65 + destcol)
                    destcol = destcol + 1
                    
                Next j
                
                destcol = 9
                count = count + 1
                destrow = destrow + 1
                
            ElseIf rd.Cells(i, rd_start_col) <> rd.Cells(i + 1, rd_start_col) Then
                
                If IsEmpty(rd.Cells(i + 1, rd_start_col)) Then Exit For
                
                count = 0
            
            End If
            
        Next i
    
    std.Cells(1, 9) = "NFH"

End Sub


Public Sub FMM_by_type_data()

    'FMM______________________________________________________________________________
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    crh_lastrow = rd.Range("CE" & rd.Rows.count).End(xlUp).Row
    Dim count As Integer
    
    count = 0
    destrow = 2
    rd_start_col = 83
    destcol = 13
        
        For i = 2 To crh_lastrow
            
            If rd.Cells(i, rd_start_col) = rd.Cells(i + 1, rd_start_col) And count < 5 Then
                
                For j = 1 To 3
                    
                    std.Cells(destrow, destcol) = rd.Cells(i, 70 + destcol)
                    destcol = destcol + 1
                    
                Next j
                
                destcol = 13
                count = count + 1
                destrow = destrow + 1
                
            ElseIf rd.Cells(i, rd_start_col) <> rd.Cells(i + 1, rd_start_col) Then
                
                If IsEmpty(rd.Cells(i + 1, rd_start_col)) Then Exit For
                
                count = 0
            
            End If
            
        Next i
    
    std.Cells(1, 13) = "FMM"

End Sub


Public Sub GHM_by_type_data()

    'GHM______________________________________________________________________________
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    crh_lastrow = rd.Range("CN" & rd.Rows.count).End(xlUp).Row
    Dim count As Integer
    
    count = 0
    destrow = 2
    rd_start_col = 92
    destcol = 17
        
        For i = 2 To crh_lastrow
            
            If rd.Cells(i, rd_start_col) = rd.Cells(i + 1, rd_start_col) And count < 5 Then
                
                For j = 1 To 3
                    
                    std.Cells(destrow, destcol) = rd.Cells(i, 75 + destcol)
                    destcol = destcol + 1
                    
                Next j
                
                destcol = 17
                count = count + 1
                destrow = destrow + 1
                
            ElseIf rd.Cells(i, rd_start_col) <> rd.Cells(i + 1, rd_start_col) Then
                
                If IsEmpty(rd.Cells(i + 1, rd_start_col)) Then Exit For
                
                count = 0
            
            End If
            
        Next i
    
    std.Cells(1, destcol) = "GHM"

End Sub

Public Sub PKM_by_type_data()

    'PKM______________________________________________________________________________
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    crh_lastrow = rd.Range("CW" & rd.Rows.count).End(xlUp).Row
    Dim count As Integer
    
    count = 0
    destrow = 2
    rd_start_col = 101
    destcol = 21
        
        For i = 2 To crh_lastrow
            
            If rd.Cells(i, rd_start_col) = rd.Cells(i + 1, rd_start_col) And count < 5 Then
                
                For j = 1 To 3
                    
                    std.Cells(destrow, destcol) = rd.Cells(i, 80 + destcol)
                    destcol = destcol + 1
                    
                Next j
                
                destcol = 21
                count = count + 1
                destrow = destrow + 1
                
            ElseIf rd.Cells(i, rd_start_col) <> rd.Cells(i + 1, rd_start_col) Then
                
                If IsEmpty(rd.Cells(i + 1, rd_start_col)) Then Exit For
                
                count = 0
            
            End If
            
        Next i
    
    std.Cells(1, destcol) = "PKM"


End Sub


Public Sub KTE_by_type_data()

    'KTE______________________________________________________________________________
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    crh_lastrow = rd.Range("DB" & rd.Rows.count).End(xlUp).Row
    Dim count As Integer
    
    count = 0
    destrow = 2
    rd_start_col = 110
    destcol = 25
        
        For i = 2 To crh_lastrow
            
            If rd.Cells(i, rd_start_col) = rd.Cells(i + 1, rd_start_col) And count < 5 Then
                
                For j = 1 To 3
                    
                    std.Cells(destrow, destcol) = rd.Cells(i, 85 + destcol)
                    destcol = destcol + 1
                    
                Next j
                
                destcol = 25
                count = count + 1
                destrow = destrow + 1
                
            ElseIf rd.Cells(i, rd_start_col) <> rd.Cells(i + 1, rd_start_col) Then
                
                If IsEmpty(rd.Cells(i + 1, rd_start_col)) Then Exit For
                
                count = 0
            
            End If
            
        Next i
    
    std.Cells(1, destcol) = "KTE"


End Sub


Public Sub RVE_by_type_data()

    'RVE______________________________________________________________________________
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    crh_lastrow = rd.Range("DO" & rd.Rows.count).End(xlUp).Row
    Dim count As Integer
    
    count = 0
    destrow = 2
    rd_start_col = 119
    destcol = 29
        
        For i = 2 To crh_lastrow
            
            If rd.Cells(i, rd_start_col) = rd.Cells(i + 1, rd_start_col) And count < 5 Then
                
                For j = 1 To 3
                    
                    std.Cells(destrow, destcol) = rd.Cells(i, 90 + destcol)
                    destcol = destcol + 1
                    
                Next j
                
                destcol = 29
                count = count + 1
                destrow = destrow + 1
                
            ElseIf rd.Cells(i, rd_start_col) <> rd.Cells(i + 1, rd_start_col) Then
                
                If IsEmpty(rd.Cells(i + 1, rd_start_col)) Then Exit For
                
                count = 0
            
            End If
            
        Next i
    
    std.Cells(1, destcol) = "RVE"


End Sub

Public Sub FME_by_type_data()

    'FME______________________________________________________________________________
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim rd As Worksheet: Set rd = ThisWorkbook.Worksheets("raw_data")
    crh_lastrow = rd.Range("DX" & rd.Rows.count).End(xlUp).Row
    Dim count As Integer
    
    count = 0
    destrow = 2
    rd_start_col = 128
    destcol = 33
        
        For i = 2 To crh_lastrow
            
            If rd.Cells(i, rd_start_col) = rd.Cells(i + 1, rd_start_col) And count < 5 Then
                
                For j = 1 To 3
                    
                    std.Cells(destrow, destcol) = rd.Cells(i, 95 + destcol)
                    destcol = destcol + 1
                    
                Next j
                
                destcol = 33
                count = count + 1
                destrow = destrow + 1
                
            ElseIf rd.Cells(i, rd_start_col) <> rd.Cells(i + 1, rd_start_col) Then
                
                If IsEmpty(rd.Cells(i + 1, rd_start_col)) Then Exit For
                
                count = 0
            
            End If
            
        Next i
    
    std.Cells(1, destcol) = "FME"


End Sub


Public Sub std_delete_cell_contents()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim count As Integer
    Dim rng As Range
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'crh
    lastRow = std.Range("A" & std.Rows.count).End(xlUp).Row
    count = 1
        
    For i = 2 To lastRow
        
        If count = 3 Or _
           count = 2 Or _
           count = 4 Or _
           count = 5 Then
                
                std.Cells(i, 1).ClearContents
        
        Else
        End If
        
        count = count + 1
        
        If count > 5 Then count = 1
    
    Next i
    
    'fmh
    lastRow = std.Range("E" & std.Rows.count).End(xlUp).Row
    count = 1
        
    For i = 2 To lastRow
        
        If count = 3 Or _
           count = 2 Or _
           count = 4 Or _
           count = 5 Then
                
                std.Cells(i, 5).ClearContents
        
        Else
        End If
        
        count = count + 1
        
        If count > 5 Then count = 1
    
    Next i
    
    'nfh
    lastRow = std.Range("I" & std.Rows.count).End(xlUp).Row
    count = 1
        
    For i = 2 To lastRow
        
        If count = 3 Or _
           count = 2 Or _
           count = 4 Or _
           count = 5 Then
                
                std.Cells(i, 9).ClearContents
        
        Else
        End If
        
        count = count + 1
        
        If count > 5 Then count = 1
    
    Next i

    'fmm
    lastRow = std.Range("M" & std.Rows.count).End(xlUp).Row
    count = 1
        
    For i = 2 To lastRow
        
        If count = 3 Or _
           count = 2 Or _
           count = 4 Or _
           count = 5 Then
                
                std.Cells(i, 13).ClearContents
        
        Else
        End If
        
        count = count + 1
        
        If count > 5 Then count = 1
    
    Next i

    'GHM
    lastRow = std.Range("Q" & std.Rows.count).End(xlUp).Row
    count = 1
        
    For i = 2 To lastRow
        
        If count = 3 Or _
           count = 2 Or _
           count = 4 Or _
           count = 5 Then
                
                std.Cells(i, 17).ClearContents
        
        Else
        End If
        
        count = count + 1
        
        If count > 5 Then count = 1
    
    Next i
    
    'PKM
    lastRow = std.Range("U" & std.Rows.count).End(xlUp).Row
    count = 1
        
    For i = 2 To lastRow
        
        If count = 3 Or _
           count = 2 Or _
           count = 4 Or _
           count = 5 Then
                
                std.Cells(i, 21).ClearContents
        
        Else
        End If
        
        count = count + 1
        
        If count > 5 Then count = 1
    
    Next i

    'KTE
    lastRow = std.Range("Y" & std.Rows.count).End(xlUp).Row
    count = 1
        
    For i = 2 To lastRow
        
        If count = 3 Or _
           count = 2 Or _
           count = 4 Or _
           count = 5 Then
                
                std.Cells(i, 25).ClearContents
        
        Else
        End If
        
        count = count + 1
        
        If count > 5 Then count = 1
    
    Next i

    'RVE
    lastRow = std.Range("AC" & std.Rows.count).End(xlUp).Row
    count = 1
        
    For i = 2 To lastRow
        
        If count = 3 Or _
           count = 2 Or _
           count = 4 Or _
           count = 5 Then
                
                std.Cells(i, 29).ClearContents
        
        Else
        End If
        
        count = count + 1
        
        If count > 5 Then count = 1
    
    Next i

    'FME
    lastRow = std.Range("AG" & std.Rows.count).End(xlUp).Row
    count = 1
        
    For i = 2 To lastRow
        
        If count = 3 Or _
           count = 2 Or _
           count = 4 Or _
           count = 5 Then
                
                std.Cells(i, 33).ClearContents
        
        Else
        End If
        
        count = count + 1
        
        If count > 5 Then count = 1
    
    Next i

End Sub

Public Sub sort_year_std()
    
    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    
    'crh
    lastRow = std.Cells(std.Rows.count, "A").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With std.Sort
        .SortFields.Clear
        .SortFields.Add key:=std.Range("A2:A" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=std.Range("C2:C" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange std.Range("A1:C" & lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    'FMH
    lastRow = std.Cells(std.Rows.count, "E").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With std.Sort
        .SortFields.Clear
        .SortFields.Add key:=std.Range("E2:E" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=std.Range("G2:G" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange std.Range("E1:G" & lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    'NFH
    lastRow = std.Cells(std.Rows.count, "I").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With std.Sort
        .SortFields.Clear
        .SortFields.Add key:=std.Range("I2:I" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=std.Range("K2:K" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange std.Range("I1:K" & lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    'FMM
    lastRow = std.Cells(std.Rows.count, "M").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With std.Sort
        .SortFields.Clear
        .SortFields.Add key:=std.Range("M2:M" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=std.Range("O2:O" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange std.Range("M1:O" & lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    'GHM
    lastRow = std.Cells(std.Rows.count, "Q").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With std.Sort
        .SortFields.Clear
        .SortFields.Add key:=std.Range("Q2:Q" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=std.Range("S2:S" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange std.Range("Q1:S" & lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    'PKM
    lastRow = std.Cells(std.Rows.count, "U").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With std.Sort
        .SortFields.Clear
        .SortFields.Add key:=std.Range("U2:U" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=std.Range("W2:W" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange std.Range("U1:W" & lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    'KTE
    lastRow = std.Cells(std.Rows.count, "Y").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With std.Sort
        .SortFields.Clear
        .SortFields.Add key:=std.Range("Y2:Y" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=std.Range("AA2:AA" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange std.Range("Y1:AA" & lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    'RVE
    lastRow = std.Cells(std.Rows.count, "AC").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With std.Sort
        .SortFields.Clear
        .SortFields.Add key:=std.Range("AC2:AC" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=std.Range("AE2:AE" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange std.Range("AC1:AE" & lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    'FME
    lastRow = std.Cells(std.Rows.count, "AG").End(xlUp).Row

    ' Sort the range based on values in column C in descending order
    With std.Sort
        .SortFields.Clear
        .SortFields.Add key:=std.Range("AG2:AG" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SortFields.Add key:=std.Range("AI2:AI" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange std.Range("AG1:AI" & lastRow)
        .Header = xlYes ' Assumes first row is header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    

End Sub



Public Sub CRH_by_year_by_type_chart()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartRange As Range
    lastRow = std.Range("B" & std.Rows.count).End(xlUp).Row
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'CRH
    targetCol = 1
    targetRow = 2
    
    Set chartRange = std.Range(std.Cells(targetRow, targetCol), std.Cells(lastRow, targetCol + 2))
    
    Set chartObj = sbtby.ChartObjects.Add(Left:=0, Width:=375, Top:=0, Height:=450)
            
        With chartObj.Chart
            .ChartType = xlBarClustered
            .ChartStyle = 222
            .SetSourceData Source:=chartRange
            .HasTitle = True
            .ChartTitle.Text = "CRH - Tickets Closed by Year by Type"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Descriptions & Years"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
            .HasLegend = False
            
            Dim s As Series
            For Each s In .SeriesCollection
                s.ApplyDataLabels
            Next s
        
        End With
    
End Sub
    
Public Sub FMH_by_year_by_type_chart()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartRange As Range
    lastRow = std.Range("F" & std.Rows.count).End(xlUp).Row
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'FMH
    targetCol = 5
    targetRow = 2
    
    Set chartRange = std.Range(std.Cells(targetRow, targetCol), std.Cells(lastRow, targetCol + 2))
    
    Set chartObj = sbtby.ChartObjects.Add(Left:=390, Width:=375, Top:=0, Height:=450)
            
        With chartObj.Chart
            .ChartType = xlBarClustered
            .ChartStyle = 222
            .SetSourceData Source:=chartRange
            .HasTitle = True
            .ChartTitle.Text = "FMH - Tickets Closed by Year by Type"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Descriptions & Years"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
            .HasLegend = False
            
            Dim s As Series
            For Each s In .SeriesCollection
                s.ApplyDataLabels
            Next s
        
        End With
    
End Sub
    
Public Sub NFH_by_year_by_type_chart()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartRange As Range
    lastRow = std.Range("j" & std.Rows.count).End(xlUp).Row
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'NFH
    targetCol = 9
    targetRow = 2
    
    Set chartRange = std.Range(std.Cells(targetRow, targetCol), std.Cells(lastRow, targetCol + 2))
    
    Set chartObj = sbtby.ChartObjects.Add(Left:=780, Width:=375, Top:=0, Height:=450)
            
        With chartObj.Chart
            .ChartType = xlBarClustered
            .ChartStyle = 222
            .SetSourceData Source:=chartRange
            .HasTitle = True
            .ChartTitle.Text = "NFH - Tickets Closed by Year by Type"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Descriptions & Years"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
            .HasLegend = False
            
            Dim s As Series
            For Each s In .SeriesCollection
                s.ApplyDataLabels
            Next s
        
        End With
    
End Sub

Public Sub FMM_by_year_by_type_chart()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartRange As Range
    lastRow = std.Range("N" & std.Rows.count).End(xlUp).Row
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'FMM
    targetCol = 13
    targetRow = 2
    
    Set chartRange = std.Range(std.Cells(targetRow, targetCol), std.Cells(lastRow, targetCol + 2))
    
    Set chartObj = sbtby.ChartObjects.Add(Left:=0, Width:=375, Top:=465, Height:=450)
            
        With chartObj.Chart
            .ChartType = xlBarClustered
            .ChartStyle = 222
            .SetSourceData Source:=chartRange
            .HasTitle = True
            .ChartTitle.Text = "FMM - Tickets Closed by Year by Type"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Descriptions & Years"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
            .HasLegend = False
            
            Dim s As Series
            For Each s In .SeriesCollection
                s.ApplyDataLabels
            Next s
        
        End With
    
End Sub


Public Sub GHM_by_year_by_type_chart()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartRange As Range
    lastRow = std.Range("R" & std.Rows.count).End(xlUp).Row
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'GHM
    targetCol = 17
    targetRow = 2
    
    Set chartRange = std.Range(std.Cells(targetRow, targetCol), std.Cells(lastRow, targetCol + 2))
    
    Set chartObj = sbtby.ChartObjects.Add(Left:=390, Width:=375, Top:=465, Height:=450)
            
        With chartObj.Chart
            .ChartType = xlBarClustered
            .ChartStyle = 222
            .SetSourceData Source:=chartRange
            .HasTitle = True
            .ChartTitle.Text = "GHM - Tickets Closed by Year by Type"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Descriptions & Years"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
            .HasLegend = False
            
            Dim s As Series
            For Each s In .SeriesCollection
                s.ApplyDataLabels
            Next s
        
        End With
    
End Sub

Public Sub PKM_by_year_by_type_chart()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartRange As Range
    lastRow = std.Range("R" & std.Rows.count).End(xlUp).Row
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'PKM
    targetCol = 21
    targetRow = 2
    
    Set chartRange = std.Range(std.Cells(targetRow, targetCol), std.Cells(lastRow, targetCol + 2))
    
    Set chartObj = sbtby.ChartObjects.Add(Left:=780, Width:=375, Top:=465, Height:=450)
            
        With chartObj.Chart
            .ChartType = xlBarClustered
            .ChartStyle = 222
            .SetSourceData Source:=chartRange
            .HasTitle = True
            .ChartTitle.Text = "PKM - Tickets Closed by Year by Type"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Descriptions & Years"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
            .HasLegend = False
            
            Dim s As Series
            For Each s In .SeriesCollection
                s.ApplyDataLabels
            Next s
        
        End With
    
End Sub

Public Sub KTE_by_year_by_type_chart()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartRange As Range
    lastRow = std.Range("Z" & std.Rows.count).End(xlUp).Row
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'KTE
    targetCol = 25
    targetRow = 2
    
    Set chartRange = std.Range(std.Cells(targetRow, targetCol), std.Cells(lastRow, targetCol + 2))
    
    Set chartObj = sbtby.ChartObjects.Add(Left:=0, Width:=375, Top:=930, Height:=450)
            
        With chartObj.Chart
            .ChartType = xlBarClustered
            .ChartStyle = 222
            .SetSourceData Source:=chartRange
            .HasTitle = True
            .ChartTitle.Text = "KTE - Tickets Closed by Year by Type"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Descriptions & Years"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
            .HasLegend = False
            
            Dim s As Series
            For Each s In .SeriesCollection
                s.ApplyDataLabels
            Next s
        
        End With
    
End Sub

Public Sub RVE_by_year_by_type_chart()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartRange As Range
    lastRow = std.Range("AD" & std.Rows.count).End(xlUp).Row
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'RVE
    targetCol = 29
    targetRow = 2
    
    Set chartRange = std.Range(std.Cells(targetRow, targetCol), std.Cells(lastRow, targetCol + 2))
    
    Set chartObj = sbtby.ChartObjects.Add(Left:=390, Width:=375, Top:=930, Height:=450)
            
        With chartObj.Chart
            .ChartType = xlBarClustered
            .ChartStyle = 222
            .SetSourceData Source:=chartRange
            .HasTitle = True
            .ChartTitle.Text = "RVE - Tickets Closed by Year by Type"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Descriptions & Years"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
            .HasLegend = False
            
            Dim s As Series
            For Each s In .SeriesCollection
                s.ApplyDataLabels
            Next s
        
        End With
    
End Sub

Public Sub FME_by_year_by_type_chart()

    Dim std As Worksheet: Set std = ThisWorkbook.Worksheets("school_type_data")
    Dim sbtby As Worksheet: Set sbtby = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartRange As Range
    lastRow = std.Range("AH" & std.Rows.count).End(xlUp).Row
    Dim targetCol As Integer
    Dim targetRow As Integer
    
    
    'FME
    targetCol = 33
    targetRow = 2
    
    Set chartRange = std.Range(std.Cells(targetRow, targetCol), std.Cells(lastRow, targetCol + 2))
    
    Set chartObj = sbtby.ChartObjects.Add(Left:=780, Width:=375, Top:=930, Height:=450)
            
        With chartObj.Chart
            .ChartType = xlBarClustered
            .ChartStyle = 222
            .SetSourceData Source:=chartRange
            .HasTitle = True
            .ChartTitle.Text = "FME - Tickets Closed by Year by Type"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Descriptions & Years"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "Tickets Closed"
            .HasLegend = False
            
            Dim s As Series
            For Each s In .SeriesCollection
                s.ApplyDataLabels
            Next s
        
        End With
    
End Sub

Sub ColorBarsByDescriptionAcrossCharts()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Schools_by_type_by_year")
    Dim chartObj As ChartObject
    Dim chartSeries As Series
    Dim description As String
    Dim colorDict As Object
    Dim colorIndex As Integer
    Dim pointIndex As Integer
    Dim color As Long
    Dim colorPalette As Variant

    ' Define a color palette with at least 30 colors
    colorPalette = Array(RGB(255, 0, 0), RGB(0, 255, 0), RGB(0, 0, 255), _
                         RGB(255, 255, 0), RGB(255, 0, 255), RGB(0, 255, 255), _
                         RGB(128, 0, 0), RGB(0, 128, 0), RGB(0, 0, 128), _
                         RGB(128, 128, 0), RGB(128, 0, 128), RGB(0, 128, 128), _
                         RGB(192, 192, 192), RGB(255, 165, 0), RGB(255, 192, 203), _
                         RGB(75, 0, 130), RGB(238, 130, 238), RGB(144, 238, 144), _
                         RGB(210, 105, 30), RGB(173, 216, 230), RGB(255, 69, 0), _
                         RGB(154, 205, 50), RGB(139, 69, 19), RGB(255, 105, 180), _
                         RGB(64, 224, 208), RGB(0, 100, 0), RGB(205, 92, 92), _
                         RGB(0, 0, 139), RGB(224, 255, 255), RGB(0, 191, 255), _
                         RGB(123, 104, 238), RGB(255, 228, 196), RGB(47, 79, 79))

    ' Initialize the dictionary to store colors for each description
    Set colorDict = CreateObject("Scripting.Dictionary")
    colorIndex = 0

    ' Loop through each chart on the worksheet
    For Each chartObj In ws.ChartObjects
        ' Loop through each series in the chart
        For Each chartSeries In chartObj.Chart.SeriesCollection
            ' Loop through each point in the series
            For pointIndex = 1 To chartSeries.Points.count
                ' Get the description from the XValues (assuming XValues are descriptions)
                description = chartSeries.XValues(pointIndex)

                ' Remove the first 5 characters to strip the year
                If Len(description) > 5 Then
                    description = Mid(description, 6)
                End If

                ' Check if the description already has an assigned color
                If Not colorDict.Exists(description) Then
                    ' Assign a new color for this description
                    colorDict(description) = colorPalette(colorIndex Mod UBound(colorPalette) + 1)
                    colorIndex = colorIndex + 1
                End If

                ' Retrieve the existing color for the description
                color = colorDict(description)

                ' Apply the color to the current point
                chartSeries.Points(pointIndex).Format.Fill.ForeColor.RGB = color
            Next pointIndex
        Next chartSeries
    Next chartObj
End Sub









'TOOLS**********************************************************************************
Public Sub toggle_R1C1()

    ' Check the current reference style
    If Application.ReferenceStyle = xlA1 Then
        ' If current style is A1, switch to R1C1
        Application.ReferenceStyle = xlR1C1
        'MsgBox "Reference style changed to R1C1."
    Else
        ' If current style is R1C1, switch to A1
        Application.ReferenceStyle = xlA1
        'MsgBox "Reference style changed to A1."
    End If

End Sub


```
