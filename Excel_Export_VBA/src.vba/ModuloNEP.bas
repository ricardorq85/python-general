Option Explicit


Public Function nep(nmEmb As String, fecha As Date) As Single
    Dim nepMesAct As Single
    Dim nepMesSig As Single
    Dim mesAct As Integer
    Dim mesSig As Integer
    Dim m As Single  'pendiente
    Dim nroDias As Integer
    Dim dia As Integer
    
    nroDias = nroDiasMes(fecha)
    mesAct = Month(fecha)
    If mesAct = 12 Then
        mesSig = 1
    Else
        mesSig = mesAct + 1
    End If
    dia = Day(fecha)
    nepMesAct = NepMes(nmEmb, mesAct)
    nepMesSig = NepMes(nmEmb, mesSig)
    m = (nepMesSig - nepMesAct) / nroDias
    nep = nepMesAct + m * (dia - 1)

End Function

Public Function NepMes(nmEmb As String, mes As Integer) As Single
    
    Dim strEmbalse As String
    Dim blnHallado As Boolean
    Dim fila As Integer
    
    nmEmb = UCase(Trim(nmEmb))
    blnHallado = False
    fila = 1
    strEmbalse = UCase(Trim(ThisWorkbook.Worksheets("NEP").Cells(fila, 1).Value))
    
    Do While blnHallado = False And fila < MAXEMBALSES
        DoEvents
        If strEmbalse = nmEmb Then
            blnHallado = True
            NepMes = ThisWorkbook.Worksheets("NEP").Cells(fila, mes + 1).Value
        End If
    
        fila = fila + 1
        strEmbalse = UCase(Trim(ThisWorkbook.Worksheets("NEP").Cells(fila, 1).Value))
    Loop

End Function

Public Function nroDiasMes(fecha As Date) As Integer
    Dim mes As Integer
    Dim año As Integer
    
    mes = Month(fecha)
    año = Year(fecha)
    
    Select Case mes
        Case 1, 3, 5, 7, 8, 10, 12
            nroDiasMes = 31
        Case 4, 6, 9, 11
            nroDiasMes = 30
        Case 2
            If año Mod 4 = 0 Then
                nroDiasMes = 29
            Else
                nroDiasMes = 28
            End If
            
    End Select


End Function