Option Explicit

Private Sub Chart_BeforeDoubleClick(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long, Cancel As Boolean)
    
    Dim myX As Variant
    Cancel = True
    
    With ActiveChart
        ' Did we click over a point or data label?
        If ElementID = xlSeries Then
            If Arg2 > 0 Then
                ' Extract x value from array of x values
                myX = WorksheetFunction.Index _
                    (.SeriesCollection(Arg1).XValues, Arg2)
                On Error Resume Next
                ' Activate the appropriate chart
                Call Macro1(myX, .SeriesCollection(Arg1).Name)
                On Error GoTo 0
            End If
        End If
    End With

End Sub
'*********************************************************************************************************************************
Sub Macro1(ByVal divisionField, ByVal dateField)
    Dim x, y As Integer
    Dim sheetName, xtableName As String
    y = 0

    Sheets.Add After:=Sheets(Sheets.Count)

    For x = 1 To Sheets.Count
        If InStr(1, Sheets(x).Name, "Drill") Then y = y + 1
    Next
    sheetName = "Drill" & y
    xtableName = "Drill Down" & y
    Sheets(Sheets.Count).Name = sheetName

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table1[#All]", Version:=xlPivotTableVersion15).CreatePivotTable TableDestination _
        :=sheetName & "!R1C1", tableName:=xtableName, DefaultVersion:=xlPivotTableVersion15

    With Sheets(sheetName).PivotTables(xtableName)
        With .PivotFields("_YEAR")
            .Orientation = xlPageField
            .Position = 1
            .CurrentPage = dateField
            .Caption = "YEAR"
        End With
        With .PivotFields("DIVISION_NAME")
            .Orientation = xlPageField
            .Position = 2
            .CurrentPage = divisionField
            .Caption = "DIVISION NAME"
        End With
        With .PivotFields("BRANCH_NAME")
            .Orientation = xlRowField
            .Position = 1 '
        End With
        With .PivotFields("LINE_COUNT_OF_ORDERS")
            .Orientation = xlDataField
            .Position = 1
            .Caption = "LINE COUNT OF ORDERS"
        End With
        With .PivotFields("INVOICE_AMOUNT")
            .Orientation = xlDataField
            .Position = 2
            .NumberFormat = "$#,##0.00"
            .Caption = "INVOICE AMOUNT"
        End With
        With .PivotFields("_QUANTITY")
            .Orientation = xlDataField
            .Position = 3
            .Caption = "QUANTITY"
        End With
        With .PivotFields("ASSOC_AS_%_OF_TYPE_1")
            .Orientation = xlDataField
            .Position = 4
            .NumberFormat = "0.0000%"
            .Caption = "ASSOC AS % OF TYPE 1"
        End With
    End With
    Sheets(sheetName).PivotTables(xtableName).CompactLayoutRowHeader = "BRANCH NAME"
End Sub
