Public Str As Long, Stlb As Long

Sub Calendar1()
    Str = ActiveCell.Row
    Stlb = ActiveCell.Column
    Calendar.Show vbModeless
End Sub