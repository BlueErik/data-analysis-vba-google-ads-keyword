Attribute VB_Name = "Module1"
Sub Search_Report()
    Dim wsKW As Worksheet, wsTerm As Worksheet, wsKW_New As Worksheet, wsTerm_New As Worksheet
    Dim ptKW_1 As PivotTable, ptKW_2 As PivotTable, ptTerm_1 As PivotTable, ptTerm_2 As PivotTable
    Dim pcKW_1 As PivotCache, pcKW_2 As PivotCache, pcTerm_1 As PivotCache, pcTerm_2 As PivotCache
    Dim dataKW As Range, dataTerm As Range
    Dim lastRowKW As Long, lastColKW As Long, lastRowTerm As Long, lastColTerm As Long, r As Long, i As Long, j As Long, k As Long
    Dim pivotStartCellKW_1 As Range, pivotStartCellKW_2 As Range, pivotStartCellTerm_1 As Range, pivotStartCellTerm_2 As Range
    Dim rngKW_1 As Range, rngKW_2 As Range, rngTerm_1 As Range, rngTerm_2 As Range
    
    BenchMark = Timer
    Application.ScreenUpdating = False
    
    ' Set up the input sheets
    Set wsKW = ThisWorkbook.Sheets("Raw-Search keyword report")
    Set wsTerm = ThisWorkbook.Sheets("Raw-Search terms report")
    
    wsKW.AutoFilterMode = False
    wsTerm.AutoFilterMode = False

    ' Set up the threshold values
    CTR_threshold = 0.0517 ' 5.17%
    CPC_threshold = 0.97 ' $0.97
    
    ' Create the output sheets
    Call SubFunc_CreateNewSheet("Search KW Report", wsKW_New)
    Call SubFunc_CreateNewSheet("Search Term Report", wsTerm_New)
    
    ' Write some basic titles
    wsKW_New.Range("A1").Value = "Keyword Performance"
    wsKW_New.Range("J1").Value = "Ad Group Keyword Performance"
    
    wsTerm_New.Range("A1").Value = "Search Term Performance"
    wsTerm_New.Range("H1").Value = "Search Term Keyword Performance"
    
    ' Define data range for Search Keyword
    lastRowKW = wsKW.Cells(Rows.Count, 1).End(xlUp).Row
    ' Data start from row 3
    lastColKW = wsKW.Cells(3, Columns.Count).End(xlToLeft).Column
    Set dataKW = wsKW.Range(wsKW.Cells(3, 1), wsKW.Cells(lastRowKW, lastColKW))

    ' Define data range for Search Term
    lastRowTerm = wsTerm.Cells(Rows.Count, 1).End(xlUp).Row
    ' Data start from row 3
    lastColTerm = wsTerm.Cells(3, Columns.Count).End(xlToLeft).Column
    Set dataTerm = wsTerm.Range(wsTerm.Cells(3, 1), wsTerm.Cells(lastRowTerm, lastColTerm))
    
    ' Define PivotTable start positions
    Set pivotStartCellKW_1 = wsKW_New.Range("A3")
    Set pivotStartCellKW_2 = wsKW_New.Range("J3")
    
    Set pivotStartCellTerm_1 = wsTerm_New.Range("A3")
    Set pivotStartCellTerm_2 = wsTerm_New.Range("H3")
    
    ' Create PivotTable
    Set pcKW_1 = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataKW)
    Set ptKW_1 = pcKW_1.CreatePivotTable(TableDestination:=pivotStartCellKW_1, TableName:="PivotKW_1")

    Set pcKW_2 = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataKW)
    Set ptKW_2 = pcKW_2.CreatePivotTable(TableDestination:=pivotStartCellKW_2, TableName:="PivotKW_2")
    
    Set pcTerm_1 = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataTerm)
    Set ptTerm_1 = pcTerm_1.CreatePivotTable(TableDestination:=pivotStartCellTerm_1, TableName:="PivotTerm_1")

    Set pcTerm_2 = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataTerm)
    Set ptTerm_2 = pcTerm_2.CreatePivotTable(TableDestination:=pivotStartCellTerm_2, TableName:="PivotTerm_2")

    ' Create the first pivot table in the keyword sheet
    DoEvents
    With ptKW_1
        ' Add the Calculated Field for CTR with IFERROR logic
        .CalculatedFields.Add Name:="CTR_Calc", Formula:="=IFERROR(Clicks/Impr., 0)"
        ' Add the Calculated Field for CPC with IFERROR logic
        .CalculatedFields.Add Name:="CPC_Calc", Formula:="=IFERROR(Cost/Clicks, 0)"
        
        With .PivotFields("Campaign")
             .Orientation = xlRowField
             .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
             .Caption = "Campaign"
        End With
        With .PivotFields("Ad group")
             .Orientation = xlRowField
             .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
        With .PivotFields("Keyword")
             .Orientation = xlRowField
             .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
        With .PivotFields("Cost")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "$#,##0.00"
        End With
        With .PivotFields("Impr.")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "#,##0"
        End With
        With .PivotFields("Clicks")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "#,##0"
        End With
        With .PivotFields("CTR_Calc")
             .Orientation = xlDataField
             .NumberFormat = "0.00%"
        End With
        With .PivotFields("CPC_Calc")
             .Orientation = xlDataField
             .NumberFormat = "$#,##0.00"
        End With
    .RowAxisLayout xlTabularRow
    .TableStyle2 = "PivotStyleLight16"
    End With
    Call SubFunc_PivotRepeatLabel(ptKW_1, False)

    ' Create the second pivot table in the keyword sheet
    With ptKW_2
        ' Add the Calculated Field for CTR with IFERROR logic
        .CalculatedFields.Add Name:="CTR_Calc", Formula:="=IFERROR(Clicks/Impr., 0)"
        ' Add the Calculated Field for CPC with IFERROR logic
        .CalculatedFields.Add Name:="CPC_Calc", Formula:="=IFERROR(Cost/Clicks, 0)"
        
        With .PivotFields("Campaign")
             .Orientation = xlRowField
             .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
             .Caption = "Campaign"
        End With
        With .PivotFields("Ad group")
             .Orientation = xlRowField
             .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
        With .PivotFields("Cost")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "$#,##0.00"
        End With
        With .PivotFields("Impr.")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "#,##0"
        End With
        With .PivotFields("Clicks")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "#,##0"
        End With
        With .PivotFields("CTR_Calc")
             .Orientation = xlDataField
             .NumberFormat = "0.00%"
        End With
        With .PivotFields("CPC_Calc")
             .Orientation = xlDataField
             .NumberFormat = "$#,##0.00"
        End With
    .RowAxisLayout xlTabularRow
    .TableStyle2 = "PivotStyleLight20"
    End With
    Call SubFunc_PivotRepeatLabel(ptKW_2, False)
    
    ' Create the first pivot table in the search term sheet
    With ptTerm_1
        ' Add the Calculated Field for CTR with IFERROR logic
        .CalculatedFields.Add Name:="CTR_Calc", Formula:="=IFERROR(Clicks/Impr., 0)"
        ' Add the Calculated Field for CPC with IFERROR logic
        .CalculatedFields.Add Name:="CPC_Calc", Formula:="=IFERROR(Cost/Clicks, 0)"
        
        With .PivotFields("Search term")
             .Orientation = xlRowField
             .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
             .Caption = "Search term"
        End With
        With .PivotFields("Cost")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "$#,##0.00"
        End With
        With .PivotFields("Impr.")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "#,##0"
        End With
        With .PivotFields("Clicks")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "#,##0"
        End With
        With .PivotFields("CTR_Calc")
             .Orientation = xlDataField
             .NumberFormat = "0.00%"
        End With
        With .PivotFields("CPC_Calc")
             .Orientation = xlDataField
             .NumberFormat = "$#,##0.00"
        End With
    .RowAxisLayout xlTabularRow
    .TableStyle2 = "PivotStyleLight16"
    End With
    Call SubFunc_PivotRepeatLabel(ptTerm_1, False)
    
    ' Create the second pivot table in the search term sheet
    With ptTerm_2
        ' Add the Calculated Field for CTR with IFERROR logic
        .CalculatedFields.Add Name:="CTR_Calc", Formula:="=IFERROR(Clicks/Impr., 0)"
        ' Add the Calculated Field for CPC with IFERROR logic
        .CalculatedFields.Add Name:="CPC_Calc", Formula:="=IFERROR(Cost/Clicks, 0)"
        
        With .PivotFields("Search term")
             .Orientation = xlRowField
             .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
             .Caption = "Search term"
        End With
        With .PivotFields("Keyword")
             .Orientation = xlRowField
             .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
        With .PivotFields("Cost")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "$#,##0.00"
        End With
        With .PivotFields("Impr.")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "#,##0"
        End With
        With .PivotFields("Clicks")
             .Orientation = xlDataField
             .Function = xlSum
             .NumberFormat = "#,##0"
        End With
        With .PivotFields("CTR_Calc")
             .Orientation = xlDataField
             .NumberFormat = "0.00%"
        End With
        With .PivotFields("CPC_Calc")
             .Orientation = xlDataField
             .NumberFormat = "$#,##0.00"
        End With
    .RowAxisLayout xlTabularRow
    .TableStyle2 = "PivotStyleLight20"
    End With
    Call SubFunc_PivotRepeatLabel(ptTerm_2, False)
    
    ' Set the pivot table ranges
    Set rngKW_1 = ptKW_1.TableRange1
    Set rngKW_2 = ptKW_2.TableRange1

    Set rngTerm_1 = ptTerm_1.TableRange1
    Set rngTerm_2 = ptTerm_2.TableRange1

    ' In the first pivot table in the keyword sheet
    For i = 2 To rngKW_1.Rows.Count - 1 ' From row 2 to last row, except Grand Total row
        Set cellName_KW = rngKW_1.Cells(i, 3)
        Set cellName_CTR = rngKW_1.Cells(i, 7)
        Set cellName_CPC = rngKW_1.Cells(i, 8)

        cellValue_CTR = cellName_CTR.Value
        cellValue_CPC = cellName_CPC.Value

        ' Find 0 < CTR < 5.17%
        If cellValue_CTR > 0 And cellValue_CTR < CTR_threshold Then
            cellName_KW.Interior.Color = ERR_HIG
            cellName_CTR.Interior.Color = ERR_HIG
            ' Find CPC > $0.97
            If cellValue_CPC > CPC_threshold Then cellName_CPC.Interior.Color = ERR_HIG
        ' Find CPC > $0.97
        ElseIf cellValue_CPC > CPC_threshold Then
            cellName_KW.Interior.Color = ERR_HIG
            cellName_CPC.Interior.Color = ERR_HIG
        End If
    Next i

    ' In the second pivot table in the keyword sheet
    For j = 2 To rngKW_2.Rows.Count - 1 ' From row 2 to last row, except Grand Total row
        Set cellName_AG = rngKW_2.Cells(j, 2)
        Set cellName_CTR = rngKW_2.Cells(j, 6)
        Set cellName_CPC = rngKW_2.Cells(j, 7)

        cellValue_CTR = cellName_CTR.Value
        cellValue_CPC = cellName_CPC.Value

        ' Find 0 < CTR < 5.17%
        If cellValue_CTR > 0 And cellValue_CTR < CTR_threshold Then
            cellName_AG.Interior.Color = ERR_HIG
            cellName_CTR.Interior.Color = ERR_HIG
            ' Find CPC > $0.97
            If cellValue_CPC > CPC_threshold Then cellName_CPC.Interior.Color = ERR_HIG
        ' Find CPC > $0.97
        ElseIf cellValue_CPC > CPC_threshold Then
            cellName_AG.Interior.Color = ERR_HIG
            cellName_CPC.Interior.Color = ERR_HIG
        End If
    Next j

    ' In the first pivot table in the search term sheet
    For r = 2 To rngTerm_1.Rows.Count - 1 ' From row 2 to last row, except Grand Total row
        Set cellName_Term = rngTerm_1.Cells(r, 1)
        Set cellName_CTR = rngTerm_1.Cells(r, 5)
        Set cellName_CPC = rngTerm_1.Cells(r, 6)

        cellValue_CTR = cellName_CTR.Value
        cellValue_CPC = cellName_CPC.Value

        ' Find 0 < CTR < 5.17%
        If cellValue_CTR > 0 And cellValue_CTR < CTR_threshold Then
            cellName_Term.Interior.Color = ERR_HIG
            cellName_CTR.Interior.Color = ERR_HIG
            ' Find CPC > $0.97
            If cellValue_CPC > CPC_threshold Then cellName_CPC.Interior.Color = ERR_HIG
        ' Find CPC > $0.97
        ElseIf cellValue_CPC > CPC_threshold Then
            cellName_Term.Interior.Color = ERR_HIG
            cellName_CPC.Interior.Color = ERR_HIG
        End If
    Next r

    ' In the second pivot table in the search term sheet
    For k = 2 To rngTerm_2.Rows.Count - 1 ' From row 2 to last row, except Grand Total row
        Set cellName_KW = rngTerm_2.Cells(k, 2)
        Set cellName_CTR = rngTerm_2.Cells(k, 6)
        Set cellName_CPC = rngTerm_2.Cells(k, 7)

        cellValue_CTR = cellName_CTR.Value
        cellValue_CPC = cellName_CPC.Value

        ' Find 0 < CTR < 5.17%
        If cellValue_CTR > 0 And cellValue_CTR < CTR_threshold Then
            cellName_KW.Interior.Color = ERR_HIG
            cellName_CTR.Interior.Color = ERR_HIG
            ' Find CPC > $0.97
            If cellValue_CPC > CPC_threshold Then cellName_CPC.Interior.Color = ERR_HIG
        ' Find CPC > $0.97
        ElseIf cellValue_CPC > CPC_threshold Then
            cellName_KW.Interior.Color = ERR_HIG
            cellName_CPC.Interior.Color = ERR_HIG
        End If
    Next k
    
    wsKW.UsedRange.Font.Name = "Aptos Narrow"
    wsKW.UsedRange.Font.Size = 12
    wsKW.Columns.AutoFit
    
    wsTerm.UsedRange.Font.Name = "Aptos Narrow"
    wsTerm.UsedRange.Font.Size = 12
    wsTerm.Columns.AutoFit

    wsKW_New.UsedRange.Font.Name = "Aptos Narrow"
    wsKW_New.UsedRange.Font.Size = 12
    wsKW_New.Columns.AutoFit
    wsKW_New.Tab.Color = TAB_HIG
    
    wsTerm_New.UsedRange.Font.Name = "Aptos Narrow"
    wsTerm_New.UsedRange.Font.Size = 12
    wsTerm_New.Columns.AutoFit
    wsTerm_New.Tab.Color = TAB_HIG
    
    Call SubFunc_SetZoomLevel(ThisWorkbook)
    wsKW_New.Select
    totalRunTime = Round((Timer - BenchMark) / 60, 2)
    Application.ScreenUpdating = True
    MsgBox "Search Reports are ready!" & vbCrLf & _
           "Total run time: " & totalRunTime & " minutes", vbInformation, "Complete"
End Sub

Public Sub SubFunc_CreateNewSheet(sheetName As String, ws As Worksheet)
    ' Helper function to create or clear a sheet by name
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
End Sub

Public Sub SubFunc_PivotRepeatLabel(pt As PivotTable, Condition As Boolean)
    ' Helper function to Design - Report Layout - Repeat all item labels
    Dim pf As PivotField
    For Each pf In pt.RowFields
        pf.RepeatLabels = Condition
    Next pf
End Sub

Public Sub SubFunc_SetZoomLevel(wb_temp As Workbook)
    Dim ws As Worksheet
    ' Helper function to set zoom level
    Set wb = wb_temp
    ' Loop through every worksheet in the workbook
    For Each ws In wb.Worksheets
        ' Skip sheets that are hidden (xlSheetHidden or xlSheetVeryHidden)
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ActiveWindow.Zoom = 100 ' Set zoom level to 100%
        End If
    Next ws
End Sub

Public Function ERR_HIG() As Long
    ' Helper function to set the light red error cell highlight color
    ERR_HIG = RGB(255, 102, 102)
End Function

Public Function TAB_HIG() As Long
    ' Helper function to set the light yellow tab highlight color
    TAB_HIG = RGB(255, 255, 153)
End Function
