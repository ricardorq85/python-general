Option Explicit

Public Type typeDDEC
    central As String
    Total As Single
    MWh(24) As Single
End Type

Public despacho(MAXCENTRALES) As typeDDEC

Sub LeerDDecGenProg(fecha As Date, Optional nroDiaAnterior As Integer)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim i As Integer
    Dim Hora As Integer
    Dim Total As Single
    
    archivo = ArchivoDDEC(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Programado_Real").Cells(2, 4 - nroDiaAnterior).Value = fecha - nroDiaAnterior

    On Error GoTo ManejadorError
    i = 0
    
    'Carga DDEC
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 24 Then
                'Debug.Print textline
                despacho(i).central = EliminarComillas(Trim(LArray(0)))
                despacho(i).Total = 0
                For Hora = 1 To 24
                    despacho(i).MWh(Hora) = LArray(Hora)
                    despacho(i).Total = despacho(i).Total + despacho(i).MWh(Hora)
                Next Hora
                i = i + 1
            End If
        Loop
    Close #1
    
    Dim FilaProgContr As Integer
    Dim posCentralDDEC As Integer
    Dim CentralDDEC As String
    Dim CentralGenProg As String
    Dim regEquiv As Integer 'Indicador de posicion en el arreglo de equivalencias
    Dim Valor As Double
    
    'Escribe GenProgramada
    OrdenarDespachoPorNombre despacho
    LeerEquivalencias
    
    FilaProgContr = 4
    CentralGenProg = ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 1).Value
    Do While CentralGenProg <> ""
        DoEvents
        regEquiv = 1
        Valor = 0
        
        
        If CentralGenProg <> "TOTAL TERMICAS" Then
            Do While regEquiv <= nroEquiv
                DoEvents
                If Equivalencias(regEquiv).informeGenProg = CentralGenProg Then
                    CentralDDEC = Equivalencias(regEquiv).CentralDDEC
                    posCentralDDEC = HallarPosCentralEnDespacho(CentralDDEC, despacho, i)
                    If posCentralDDEC <> -1 Then
                        Valor = Valor + despacho(posCentralDDEC).Total
                        'Debug.Print CentralGenProg & " " & Equivalencias(regEquiv).CentralDDEC & " " & CStr(Valor)
                    Else
                    '    ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 3 - nroDiaAnterior).Value = "No hallada"
                    End If
                End If
                regEquiv = regEquiv + 1
            Loop
        Else
            Do While regEquiv <= nroEquiv
                DoEvents
                If UCase(Trim(Equivalencias(regEquiv).Tipo)) = "GT" Then
                    CentralDDEC = Equivalencias(regEquiv).CentralDDEC
                    posCentralDDEC = HallarPosCentralEnDespacho(CentralDDEC, despacho, i)
                    If posCentralDDEC <> -1 Then
                        Valor = Valor + despacho(posCentralDDEC).Total
                        'Debug.Print CentralGenProg & " " & Equivalencias(regEquiv).CentralDDEC & " " & CStr(Valor)
                    Else
                    '    ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 3 - nroDiaAnterior).Value = "No hallada"
                    End If
                End If
                regEquiv = regEquiv + 1
            Loop
        End If
        
        ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 4 - nroDiaAnterior).Value = Valor / 1000
            
        FilaProgContr = FilaProgContr + 1
        CentralGenProg = ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 1).Value
    Loop
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerDDecGenProg"
End Sub

Sub LeerDDecGenProgHoraria(fecha As Date, Optional nroDiaAnterior As Integer)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim i As Integer
    Dim Hora As Integer
    Dim Total As Single
    Dim UltFila As Integer
    Dim Cadena As String
    
    archivo = ArchivoDDEC(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Programado_Real").Cells(2, 4 - nroDiaAnterior).Value = fecha - nroDiaAnterior
    
    i = 4
    Cadena = ThisWorkbook.Worksheets("Programado_Real").Cells(i, 1).Value
    Do While Cadena <> ""
        DoEvents
        i = i + 1
        Cadena = ThisWorkbook.Worksheets("Programado_Real").Cells(i, 1).Value
    Loop
    
    ThisWorkbook.Worksheets("Programado_Real").Select
    ThisWorkbook.Worksheets("Programado_Real").Range(Cells(4, 6).Address, Cells(i - 1, 29)).Value = 0
    'ThisWorkbook.Worksheets("Programado_Real").Range(Cells(4, 6).Address, Cells(i - 1, 29)).Clear
    
    On Error GoTo ManejadorError
    i = 0
    
    'Carga DDEC
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 24 Then
                'Debug.Print textline
                despacho(i).central = EliminarComillas(Trim(LArray(0)))
                despacho(i).Total = 0
                For Hora = 1 To 24
                    despacho(i).MWh(Hora) = LArray(Hora)
                    despacho(i).Total = despacho(i).Total + despacho(i).MWh(Hora)
                Next Hora
                i = i + 1
            End If
        Loop
    Close #1
    
    Dim FilaProgContr As Integer
    Dim posCentralDDEC As Integer
    Dim CentralDDEC As String
    Dim CentralGenProg As String
    Dim regEquiv As Integer 'Indicador de posicion en el arreglo de equivalencias
    Dim Valor As Double
    
    'Escribe GenProgramada
    OrdenarDespachoPorNombre despacho
    LeerEquivalencias
    
    FilaProgContr = 4
    CentralGenProg = ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 1).Value
    Do While CentralGenProg <> ""
        DoEvents
        regEquiv = 1
        Valor = 0
        
        If CentralGenProg <> "TOTAL TERMICAS" Then
            Do While regEquiv <= nroEquiv
                DoEvents
                If Equivalencias(regEquiv).informeGenProg = CentralGenProg Then
                    
                    CentralDDEC = Equivalencias(regEquiv).CentralDDEC
                    posCentralDDEC = HallarPosCentralEnDespacho(CentralDDEC, despacho, i)
                    If posCentralDDEC <> -1 Then
                        For Hora = 1 To 24
                            ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 5 + Hora).Value = _
                                ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 5 + Hora).Value + despacho(posCentralDDEC).MWh(Hora)
                        Next Hora
                        'Valor = Valor + despacho(posCentralDDEC).Total
                        'Debug.Print CentralGenProg & " " & Equivalencias(regEquiv).CentralDDEC & " " & CStr(Valor)
                    Else
                    '    ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 3 - nroDiaAnterior).Value = "No hallada"
                    End If
    
                End If 'Hora
            
                regEquiv = regEquiv + 1
            Loop
        Else
            Do While regEquiv <= nroEquiv
                DoEvents
                If UCase(Trim(Equivalencias(regEquiv).Tipo)) = "GT" Then
                    
                    CentralDDEC = Equivalencias(regEquiv).CentralDDEC
                    posCentralDDEC = HallarPosCentralEnDespacho(CentralDDEC, despacho, i)
                    If posCentralDDEC <> -1 Then
                        For Hora = 1 To 24
                            ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 5 + Hora).Value = _
                                ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 5 + Hora).Value + despacho(posCentralDDEC).MWh(Hora)
                        Next Hora
                    
                        'Valor = Valor + despacho(posCentralDDEC).Total
                        'Debug.Print CentralGenProg & " " & Equivalencias(regEquiv).CentralDDEC & " " & CStr(Valor)
                    Else
                    '    ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 3 - nroDiaAnterior).Value = "No hallada"
                    End If
    
                End If
            
                regEquiv = regEquiv + 1
            Loop
        End If
        
        
        'ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 6).Value = Valor / 1000
            
        FilaProgContr = FilaProgContr + 1
        CentralGenProg = ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 1).Value
    Loop
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerDDecGenProgHoraria"
    
End Sub


Sub LeerDDecGeneracion(fecha As Date, Optional nroDiaAnterior As Integer)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim i As Integer
    Dim Hora As Integer
    Dim Total As Single
    
    archivo = ArchivoDDEC(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Generacion").Cells(1, 2).Value = fecha - nroDiaAnterior

    i = 0
    On Error GoTo ManejadorError
    'Carga DDEC
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 24 Then
                'Debug.Print textline
                despacho(i).central = EliminarComillas(Trim(LArray(0)))
                despacho(i).Total = 0
                For Hora = 1 To 24
                    despacho(i).MWh(Hora) = LArray(Hora)
                    despacho(i).Total = despacho(i).Total + despacho(i).MWh(Hora)
                Next Hora
                i = i + 1
            End If
        Loop
    Close #1
    
    Dim FilaGeneracion As Integer
    Dim posCentralDDEC As Integer
    Dim CentralGeneracion As String

    Dim Valor As Double
    
    'Escribe GenProgramada
    OrdenarDespachoPorNombre despacho

    
    FilaGeneracion = 3
    CentralGeneracion = ThisWorkbook.Worksheets("Generacion").Cells(FilaGeneracion, 1).Value
    Do While CentralGeneracion <> ""
        DoEvents
        Valor = 0
                
            posCentralDDEC = HallarPosCentralEnDespacho(CentralGeneracion, despacho, i)
            If posCentralDDEC <> -1 Then
                Valor = despacho(posCentralDDEC).Total
                'Debug.Print CentralGenProg & " " & Equivalencias(regEquiv).CentralDDEC & " " & CStr(Valor)
            Else
            '    ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 3 - nroDiaAnterior).Value = "No hallada"
            End If
        
        ThisWorkbook.Worksheets("Generacion").Cells(FilaGeneracion, 2).Value = Valor / 1000
            
        FilaGeneracion = FilaGeneracion + 1
        CentralGeneracion = ThisWorkbook.Worksheets("Generacion").Cells(FilaGeneracion, 1).Value
    Loop
    
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerDDECGeneracion"
    
End Sub



Public Function ArchivoDDEC(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamDDEC, ColParamPrefijo).Value
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoDDEC = raiz & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamDDEC, ColParamRaiz).Value
        ArchivoDDEC = raiz & año & "\" & mes & "\" & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    End If
End Function


Sub OrdenarDespachoPorNombre(despacho() As typeDDEC)
    
    Dim nroCentral As Integer
    Dim nmCentral As String

    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    nroCentral = 1
    nmCentral = despacho(nroCentral).central
    Do While nmCentral <> ""
        nroCentral = nroCentral + 1
        nmCentral = despacho(nroCentral).central
    Loop
    nroCentral = nroCentral - 1
    
    'Ordenar centrales por nombre
    For i = 1 To nroCentral - 1
        For j = i + 1 To nroCentral
            If UCase(despacho(i).central) > UCase(despacho(j).central) Then
            
                despacho(0).central = despacho(i).central
                despacho(0).Total = despacho(i).Total
                For k = 1 To 24
                    despacho(0).MWh(k) = despacho(i).MWh(k)
                Next k
                
                despacho(i).central = despacho(j).central
                despacho(i).Total = despacho(j).Total
                For k = 1 To 24
                    despacho(i).MWh(k) = despacho(j).MWh(k)
                Next k
                
                despacho(j).central = despacho(0).central
                despacho(j).Total = despacho(0).Total
                For k = 1 To 24
                    despacho(j).MWh(k) = despacho(0).MWh(k)
                Next k
                
            End If
        Next j
    Next i
    

    
End Sub

Public Function HallarPosCentralEnDespacho(nmCentral As String, despacho() As typeDDEC, nroCentrales As Integer) As Integer
    Dim centralInf As String
    Dim centralSup As String
    Dim centralMed As String
    Dim blnHallado As String
    Dim intSup As Integer
    Dim intInf As Integer
    Dim intMed As Integer
    
    
    nmCentral = UCase(Trim(nmCentral))
    
    HallarPosCentralEnDespacho = -1
    
    intSup = nroCentrales
    intInf = 1
    intMed = Int((intSup + intInf) / 2)
    centralInf = UCase(Trim(despacho(intInf).central))
    centralSup = UCase(Trim(despacho(intSup).central))
    centralMed = UCase(Trim(despacho(intMed).central))
    
    If centralInf = nmCentral Then
        HallarPosCentralEnDespacho = intInf
        Exit Function
    End If
    
    If centralSup = nmCentral Then
        HallarPosCentralEnDespacho = intSup
        Exit Function
    End If
    
    If centralMed = nmCentral Then
        HallarPosCentralEnDespacho = intMed
        Exit Function
    End If
    
    blnHallado = False
    
    Do While centralMed <> nmCentral And blnHallado = False And (intSup - intInf > 1)
        
        DoEvents
        
        If nmCentral > centralMed Then
            intInf = intMed
            centralInf = UCase(Trim(despacho(intInf).central))
        End If
        
        If nmCentral < centralMed Then
            intSup = intMed
            centralSup = UCase(Trim(despacho(intSup).central))
        End If
        
        intMed = Int((intSup + intInf) / 2)
        centralMed = UCase(Trim(despacho(intMed).central))
        
        If nmCentral = centralMed Then
            HallarPosCentralEnDespacho = intMed
            blnHallado = True
        End If
    Loop

End Function


Public Sub LeerDDEC(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim fila As Integer
    Dim Hora As Integer
    Dim Total As Single
    
    Application.Calculation = xlCalculationManual
    
    ThisWorkbook.Worksheets("DDEC").UsedRange.Delete
    
    archivo = ArchivoDDEC(fecha)
    ThisWorkbook.Worksheets("DDEC").Cells(1, 1).Value = "DDEC  " & fecha
    ThisWorkbook.Worksheets("DDEC").Cells(3, 1).Value = "Central"
    For Hora = 1 To 24
        ThisWorkbook.Worksheets("DDEC").Cells(1, Hora + 1).Value = "Hora " & CStr(Hora)
    Next Hora
    ThisWorkbook.Worksheets("DDEC").Cells(1, 26).Value = "Promedio"
    ThisWorkbook.Worksheets("DDEC").Cells(1, 27).Value = "Máximo"
    ThisWorkbook.Worksheets("DDEC").Cells(3, 26).Value = "Total"

    fila = 4
    On Error GoTo ManejadorError
    'Carga DDEC
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 24 Then
                'Debug.Print textline
                Total = 0
                ThisWorkbook.Worksheets("DDEC").Cells(fila, 1).Value = EliminarComillas(LArray(0))
                For Hora = 1 To 24
                    ThisWorkbook.Worksheets("DDEC").Cells(fila, Hora + 1).Value = LArray(Hora)
                    Total = Total + LArray(Hora)
                Next Hora
                ThisWorkbook.Worksheets("DDEC").Cells(fila, Hora + 1).Value = Total
                fila = fila + 1
            End If
        Loop
    Close #1
    
    FormatoSimpleHoja "DDEC"
    LeerDMAR fecha
    RevisarPreciosMarginales fecha, "DDEC"
    
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerDDEC"
    
End Sub



