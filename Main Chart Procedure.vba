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
