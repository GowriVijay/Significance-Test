Sub Proportions_Test1()
Col1 = Cells(Rows.Count, 1).End(xlUp).Row
col2 = Cells(Rows.Count, 2).End(xlUp).Row
col3 = Cells(Rows.Count, 3).End(xlUp).Row
col4 = Cells(Rows.Count, 4).End(xlUp).Row
minvalue = Application.WorksheetFunction.Min(Col1, col2, col3, col4)
Range("B1:B" & col2).Font.Color = vbBlack
Range("B1:B" & col2).Font.Bold = False

'Start_Row = 1
Start_Row = Sheets("Inputs").Range("B1").Value
'End_Row = Sheets("Inputs").Range("B2").Value
End_Row = minvalue


Start_Col_Metrics = Sheets("Inputs").Range("B3").Value
Start_Col_SS = Sheets("Inputs").Range("B4").Value
End_Col_Metrics = Sheets("Inputs").Range("B5").Value

Diffincols = Start_Col_SS - Start_Col_Metrics

For Col = Start_Col_Metrics To End_Col_Metrics - 1

    For Row = Start_Row To End_Row
   
        Sheets("Sheet1").Select
        n1 = Sheets("Sheet1").Cells(Row, Col + Diffincols).Value
        n2 = Sheets("Sheet1").Cells(Row, Col + 1 + Diffincols).Value
        p1 = Sheets("Sheet1").Cells(Row, Col).Value
        p2 = Sheets("Sheet1").Cells(Row, Col + 1).Value
    
    
        'Satisfying standard binomial requirement
    
        If (n1 * p1) >= 5 And (n1 * (1 - p1)) >= 5 And (n2 * p2) >= 5 And (n2 * (1 - p2)) >= 5 Then
    
            Avg_Prop = ((n1 * p1) + (n2 * p2)) / (n1 + n2)
            Inv_sample_size = (1 / n1) + (1 / n2)
            Denominator = (Avg_Prop * (1 - Avg_Prop) * Inv_sample_size) ^ (1 / 2)
            Prop_Diff = p2 - p1
            zScore = Prop_Diff / Denominator
        
            If zScore > 1.96 Then
    
                Sheets("Sheet1").Select
                Cells(Row, Col + 1).Select
                With Selection.Font
                    .Color = -11489280
                    .TintAndShade = 0
                End With
                Selection.Font.Bold = True
    
            ElseIf zScore < -1.96 Then
    
                Sheets("Sheet1").Select
                Cells(Row, Col + 1).Select
                With Selection.Font
                    .Color = -16776961
                    .TintAndShade = 0
                End With
                Selection.Font.Bold = True
    
            ElseIf zScore > 1.645 Then
    
                Sheets("Sheet1").Select
                Cells(Row, Col + 1).Select
                With Selection.Font
                    .Color = -11489280
                    .TintAndShade = 0
                End With
                Selection.Font.Bold = False
    
            ElseIf zScore < -1.645 Then
    
                Sheets("Sheet1").Select
                Cells(Row, Col + 1).Select
                With Selection.Font
                    .Color = -16776961
                    .TintAndShade = 0
                End With
                Selection.Font.Bold = False
        
            Else
    
                Sheets("Sheet1").Select
                'Cells(Row + 200, Col + 1).Value = 0
    
            End If
    
        Else
    
            Sheets("Sheet1").Select
            'Cells(Row + 200, Col + 1).Value = 0
    
        End If
    
    Next Row
    
Next Col

Sheets("Sheet1").Range("A1").Activate

End Sub
Sub resetting()
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Input")
ws.Activate
ws.Range("B1:B" & col2).Font.Color = vbBlack
ws.Range("B1:B" & col2).Font.Bold = False
End Sub

Sub SampleSize1()

Start_Row = Sheets("Inputs").Range("B1").Value
End_Row = Sheets("Inputs").Range("B2").Value
Start_Col_SS = Sheets("Inputs").Range("B4").Value
End_Col_SS = Sheets("Inputs").Range("B6").Value

Diffincols = Start_Col_SS - Start_Col_Metrics
Sheets("Sheet1").Select

For Col = Start_Col_SS To End_Col_SS

    For Row = Start_Row To End_Row
    
            Value = ActiveSheet.Cells(Row, Col).Value

            If ActiveSheet.Cells(Row, Col).Value > 100 Or ActiveSheet.Cells(Row, Col).Value = "" Then
                Cells(Row, Col).Select
    
            ElseIf ActiveSheet.Cells(Row, Col).Value > 50 Then
    
                Cells(Row, Col).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorLight2
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
   
       
            ElseIf ActiveSheet.Cells(Row, Col).Value >= 0 Then
                 Cells(Row, Col).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent2
                        .TintAndShade = 0.599993896298105
                        .PatternTintAndShade = 0
                    End With
    
            End If

    
    Next Row
    
Next Col

End Sub


Sub x()

Call Proportions_Test1
Call SampleSize1

End Sub
