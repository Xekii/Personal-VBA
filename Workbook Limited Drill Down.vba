Option Explicit

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim newSheet, x, y As Integer
    Dim sheetName, tableName, charString, dateField, divisionField As String

    Cancel = True
    On Error GoTo notPivotTable
    If Target.PivotTable.Name <> "" And Intersect(Target, Target.PivotTable.RowRange.Offset(1)) <> "" And _
       "BRANCH_NAME" = Target.PivotTable.RowFields(1) Then

        dateField = Target.PivotTable.PivotFields("YEAR").CurrentPage.Name
        divisionField = Target.PivotTable.PivotFields("DIVISION NAME").CurrentPage.Name

        Sheets.Add After:=Sheets(ActiveSheet.Index)
        newSheet = ActiveSheet.Index
        
        y = 0
        For x = 1 To Sheets.Count
            If InStr(1, Sheets(x).Name, "Drill") Then y = y + 1
        Next
        sheetName = "Drill" & y
        tableName = "Drill Down" & y
        Sheets(newSheet).Name = sheetName

        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Table2[#All]", Version:=xlPivotTableVersion15).CreatePivotTable _
            TableDestination:=sheetName & "!R1C1", tableName:=tableName, _
            DefaultVersion:=xlPivotTableVersion15

        With Sheets(newSheet).PivotTables(tableName)
            With .PivotFields("_YEAR")
                .Orientation = xlPageField
                .Position = 1
                .Caption = "YEAR"
                .CurrentPage = dateField
            End With
            With .PivotFields("DIVISION_NAME")
                .Orientation = xlPageField
                .Position = 1
                .Caption = "DIVISION NAME"
                .CurrentPage = divisionField
            End With
            With .PivotFields("BRANCH_NAME")
                .Orientation = xlPageField
                .Position = 1
                .Caption = "BRANCH NAME"
                .CurrentPage = Target.Value
            End With
            With .PivotFields("ASSOCIATED_ITEM_FLAG")
                .Orientation = xlPageField
                .Position = 1
                .Caption = "ASSOC PROD FLAG"
                .CurrentPage = "A"
            End With
            With .PivotFields("PROD_CODE")
                .Orientation = xlRowField
                .Position = 1
                .Caption = "PRODUCT"
            End With
            With .PivotFields("LINE_COUNT_OF_ORDERS")
                .Orientation = xlDataField
                .Position = 1
                .Caption = "LINE COUNT OF ORDERS"
            End With
            With .PivotFields("INVOICE_AMOUNT")
                .Orientation = xlDataField
                .Position = 2
                .Caption = "INVOICE AMOUNT"
                .NumberFormat = "$#,##0.00"
            End With
            With .PivotFields("_QUANTITY")
                .Orientation = xlDataField
                .Position = 3
                .Caption = "QUANTITY"
            End With
        End With
        Target.PivotTable.CompactLayoutRowHeader = "BRANCH"
    End If
notPivotTable:
'If Err <> 0 Then MsgBox ("There was an error")
End Sub
