Attribute VB_Name = "Module1"
Option Explicit


Sub Range6()
    With Range("A1")
        Range(.Offset(0, 1), .End(xlToRight)).Name = "ScoreNames"
        Range(.Offset(1, 0), .End(xlDown)).Name = "EmployeeNumbers"
        Range(.Offset(1, 1), .End(xlDown).End(xlToRight)).Name = "ScoreData"
    End With
    
    With Range("ScoreNames")
        .HorizontalAlignment = xlRight
        With .Font
            .Bold = True
            .ColorIndex = 3
            .Size = 16
        End With
        .EntireColumn.AutoFit
    End With
    
    With Range("EmployeeNumbers").Font
        .Italic = True
        .ColorIndex = 5
        .Size = 12
    End With
    
    With Range("ScoreData")
        .Interior.ColorIndex = 15
        .Font.Name = "Times Roman"
        .NumberFormat = "0.0"
    End With
    
    MsgBox "Formatting has been applied"
    
    Cells(1, 7) = "Average Scores for Employees"
    
    With Range("G1")
        .EntireColumn.AutoFit
        .EntireColumn.HorizontalAlignment = xlCenter
    End With
    
    
    Cells(2, 7).FormulaR1C1 = "=SUM(RC[-5]:RC[-1]) / 5"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G19"), Type:=xlFillDefault
    Range("G2:G19").Select
    MsgBox "Average formulas has been applied"
    
    
    Range("A2:G19").Sort Key1:=Range("G1"), Key2:=Range("A1")
    MsgBox "Ordering has been done"
    
End Sub

