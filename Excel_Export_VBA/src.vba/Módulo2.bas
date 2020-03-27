Sub Macro3()
'
' Macro3 Macro
'

'
    Range("F178").Select
    Selection.Locked = False
    Selection.FormulaHidden = False
End Sub
Sub Macro4()
'
' Macro4 Macro
'

'
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("G178").Select
    ActiveSheet.Unprotect
End Sub