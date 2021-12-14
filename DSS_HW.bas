Option Explicit

Sub Main()
    ThisWorkbook.Names.Add Name:="Header", RefersTo:=Range("A1:F1")
    
    
    ThisWorkbook.Names.Add Name:="Region", RefersTo:=Range("A2:A26")
    ThisWorkbook.Names.Add Name:="Rep", RefersTo:=Range("B2:B26")
    ThisWorkbook.Names.Add Name:="Items", RefersTo:=Range("C2:C26")
    ThisWorkbook.Names.Add Name:="Units", RefersTo:=Range("D2:D26")
    ThisWorkbook.Names.Add Name:="UnitCost", RefersTo:=Range("E2:E26")
    ThisWorkbook.Names.Add Name:="Total", RefersTo:=Range("F2:F26")
    
    ThisWorkbook.Names.Add Name:="AllData", RefersTo:=Range("A2:F26")
    
    With Range("Header")
        .Font.Bold = True
        .Interior.ColorIndex = 15
    End With
    
    With Range("AllData")
        .Sort Key1:=Range("Items"), Order1:=xlAscending, Key2:=Range("Total"), Order2:=xlDescending
        .HorizontalAlignment = xlRight
    End With
    
    Application.Union(Range("Region"), Range("Items"), Range("UnitCost")).Interior.ColorIndex = 34
    Application.Union(Range("Rep"), Range("Units"), Range("Total")).Interior.ColorIndex = 40

End Sub


