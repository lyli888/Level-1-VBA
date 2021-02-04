VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Sub LyLiScript()

' Extracts Unique Ticker Names

    Columns("A:A").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Columns("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Columns( _
        "I:I"), Unique:=True
        
        
' Labels Headers &

    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    

 'Counts Unique Ticker Names in Sheet
 
    
    Dim totalstock As Long

    Dim LstRw As Long, Rng As Range, List As Object
    LstRw = Cells(Rows.Count, "A").End(xlUp).Row
    Set List = CreateObject("Scripting.Dictionary")

    For Each Rng In Range("A2:A" & LstRw)
    If Not List.Exists(Rng.Value) Then List.Add Rng.Value, Nothing
    Next

    totalstock = List.Count + 1
    
    
'Calculates Yearly, Percent Change, Sum Volume for each line


    Dim x As String, y As String


    Dim test As Integer, cellrow As Long, cellcol As Integer, openrow As Long, i As Long
    
    i = 1
    cellrow = 2
    cellrow2 = cellrow + 1
    cellcol = 1
    openrow = cellrow
    sumvolume = 0
    
    
    
 'Loop to Test Ticker For Same Stock And Move Row Counters
 
    
    Do While i < totalstock
    
    test = 0
    
    
    Do While test = 0
    
    
    
    cellrow = cellrow + 1
    
    cellrow2 = cellrow2 + 1
   
    x = Cells(cellrow, cellcol)
    
    y = Cells(cellrow2, cellcol)
        
    test = StrComp(x, y)
    
    
    Loop
    
    
 'Calculates And Enters Yearly Stock Change
    
    Dim closing As Variant, opening As Variant, yearly As Variant, percent As Variant
    
    
    closing = Cells(cellrow, 6)
    
    opening = Cells(openrow, 3)
    
    yearly = closing - opening
    
    i = i + 1
    
    Cells(i, 10) = yearly
    

'Enters Yearly Stock Change Percentage

    percent = (yearly / opening) * 100
    
    Cells(i, 11) = percent
    

'Enters Stock Volume

    ActiveSheet.Range(Cells(openrow, 7), Cells(cellrow, 7)).Select

    
    Cells(i, 12) = Application.WorksheetFunction.Sum(Selection)


    
'Advances to Next Ticker
    
    sumvolume = 0

    openrow = cellrow + 1
    
    
   
    Loop
   
  
'Colors In Percent Change Column
  
    Range(Range("K2"), Range("K2").End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
        
End Sub

