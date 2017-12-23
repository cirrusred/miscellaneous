VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub sub1()

Do

'copy to cell below
ActiveCell.Select
ActiveCell.Offset(1, 0).Value = ActiveCell.Value

'jump two rows
ActiveCell.Offset(2, 0).Select

Loop Until IsEmpty(ActiveCell)


End Sub

