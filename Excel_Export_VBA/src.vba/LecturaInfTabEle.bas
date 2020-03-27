
Option Explicit

Public Type typeEmpresa
    Empresa As String
    DR As Single
    DI As Single
    MillPesos As Single
End Type

Public Type typeResultado
    Resultado As String
    Valor As Single
End Type

Public Type typePrioridad
    central As String
    Precio As Single
End Type

Public Empresas(MAXEMPRESAS) As typeEmpresa
Public PrecioBolsa(24) As Single
Public Resultados(20) As typeResultado
Public Prioridades(MAXCENTRALES) As typePrioridad






Sub LeerInfTabElePreciosOferta(fecha As Date, pr() As typePrioridad, nroDiaAnterior As Integer)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim Palabra As String
    archivo = ArchivoInfTabEle(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Ofertas").Cells(2, 5 - nroDiaAnterior).Value = fecha - nroDiaAnterior
    
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    
    Dim central As String
    Dim posCentral As Integer
    Dim FilaOferta As Integer
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            textline = DejarUnSoloEspacio(textline)
            LArray = Split(textline, " ")
            If UBound(LArray) > -1 Then
                Palabra = Trim(UCase(LArray(0)))
                If UCase(Trim(LArray(0))) = "PRIORIDADES" Then
                    Line Input #1, textline
                    textline = DejarUnSoloEspacio(textline)
                    LArray = Split(textline, ";")
                    Palabra = Trim(UCase(LArray(0)))
                    i = 0
                    Do While Palabra <> ""
                        i = i + 1
                        
                        pr(i).central = LArray(0)
                        pr(i).Precio = LArray(1)
                        'Debug.Print pr(i).Central & "    " & pr(i).Precio
                        
                        Line Input #1, textline
                        textline = DejarUnSoloEspacio(textline)
                        LArray = Split(textline, ";")
                        Palabra = Trim(UCase(LArray(0)))
                        
                    Loop
                End If
            End If
        Loop
        i = i + 1
    Close #1
    
    OrdenarPrioridadesPorNombre pr
  
    FilaOferta = 3
    central = ThisWorkbook.Worksheets("Ofertas").Cells(FilaOferta, 1).Value
    Do While central <> ""
        DoEvents
        posCentral = HallarPosCentralEnPrioridades(central, pr, i)
        ThisWorkbook.Worksheets("Ofertas").Cells(FilaOferta, 5 - nroDiaAnterior).Value = pr(posCentral).Precio
        FilaOferta = FilaOferta + 1
        central = ThisWorkbook.Worksheets("Ofertas").Cells(FilaOferta, 1).Value
    Loop
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerInfTabElePreciosOferta"
End Sub

Sub OrdenarPrioridadesPorNombre(pr() As typePrioridad)
    
    Dim nroCentral As Integer
    Dim nmCentral As String

    
    Dim i As Integer
    Dim j As Integer
    
    nroCentral = 1
    nmCentral = pr(nroCentral).central
    Do While nmCentral <> ""
        nroCentral = nroCentral + 1
        nmCentral = pr(nroCentral).central
    Loop
    nroCentral = nroCentral - 1
    
    'Ordenar centrales por nombre
    For i = 1 To nroCentral - 1
        For j = i + 1 To nroCentral
            If UCase(pr(i).central) > UCase(pr(j).central) Then
            
                pr(0).central = pr(i).central
                pr(0).Precio = pr(i).Precio
                pr(i).central = pr(j).central
                pr(i).Precio = pr(j).Precio
                pr(j).central = pr(0).central
                pr(j).Precio = pr(0).Precio
                
            End If
        Next j
    Next i
    

    
End Sub

Public Function HallarPosCentralEnPrioridades(nmCentral As String, pr() As typePrioridad, nroCentrales As Integer) As Integer
    Dim centralInf As String
    Dim centralSup As String
    Dim centralMed As String
    Dim blnHallado As String
    Dim intSup As Integer
    Dim intInf As Integer
    Dim intMed As Integer
    
    
    nmCentral = UCase(Trim(nmCentral))
    
    HallarPosCentralEnPrioridades = -1
    
    intSup = nroCentrales
    intInf = 1
    intMed = Int((intSup + intInf) / 2)
    centralInf = UCase(Trim(pr(intInf).central))
    centralSup = UCase(Trim(pr(intSup).central))
    centralMed = UCase(Trim(pr(intMed).central))
    
    If centralInf = nmCentral Then
        HallarPosCentralEnPrioridades = intInf
        Exit Function
    End If
    
    If centralSup = nmCentral Then
        HallarPosCentralEnPrioridades = intSup
        Exit Function
    End If
    
    If centralMed = nmCentral Then
        HallarPosCentralEnPrioridades = intMed
        Exit Function
    End If
    
    blnHallado = False
    
    Do While centralMed <> nmCentral And blnHallado = False And (intSup - intInf > 1)
        
        If nmCentral > centralMed Then
            intInf = intMed
            centralInf = UCase(Trim(pr(intInf).central))
        End If
        
        If nmCentral < centralMed Then
            intSup = intMed
            centralSup = UCase(Trim(pr(intSup).central))
        End If
        
        intMed = Int((intSup + intInf) / 2)
        centralMed = UCase(Trim(pr(intMed).central))
        
        If nmCentral = centralMed Then
            HallarPosCentralEnPrioridades = intMed
            blnHallado = True
        End If
    Loop

End Function

Sub LeerInfTabElePrioridades(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim Palabra As String
    archivo = ArchivoInfTabEle(fecha)
    Dim i As Integer
    Dim j As Integer
    Dim Hora As Integer
    i = 0
    
    Dim central As String
    Dim FilaOferta As Integer
    
    On Error GoTo ManejadorError
    
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            textline = DejarUnSoloEspacio(textline)
            LArray = Split(textline, " ")
            If UBound(LArray) > -1 Then
                Palabra = Trim(UCase(LArray(0)))
            If UCase(Trim(LArray(0))) = "PRIORIDADES" Then
                    Line Input #1, textline
                    textline = DejarUnSoloEspacio(textline)
                    LArray = Split(textline, ";")
                    Palabra = Trim(UCase(LArray(0)))
                    i = 0
                    Do While Palabra <> ""
                        i = i + 1
                        
                        Prioridades(i).central = LArray(0)
                        Prioridades(i).Precio = LArray(1)
                        'Debug.Print Prioridades(i).Central & "    " & Prioridades(i).Precio
                        
                        Line Input #1, textline
                        textline = DejarUnSoloEspacio(textline)
                        LArray = Split(textline, ";")
                        Palabra = Trim(UCase(LArray(0)))
                    Loop
                End If
            End If
        Loop
    Close #1
    
    FilaOferta = 3
    
    For j = 1 To i
        ThisWorkbook.Worksheets("Ofertas").Cells(FilaOferta, 10).Value = Prioridades(j).central
        ThisWorkbook.Worksheets("Ofertas").Cells(FilaOferta, 11).Value = Prioridades(j).Precio
        FilaOferta = FilaOferta + 1
    Next j
    
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " CargarSEGDES_DispCen"
End Sub

Public Function ArchivoInfTabEle(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamInfTabEle, ColParamPrefijo).Value
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoInfTabEle = raiz & prefijo & NombreDia(nmDia.Corto, fecha) & NombreMes(nmMes.Corto, fecha) & nroDia(ConCero, fecha) & ".txt"
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamInfTabEle, ColParamRaiz).Value
        ArchivoInfTabEle = raiz & año & "\" & mes & "\Liquidación\" & prefijo & NombreDia(nmDia.Corto, fecha) & NombreMes(nmMes.Corto, fecha) & nroDia(ConCero, fecha) & ".txt"
    End If
End Function


Sub LeerInfTabElePrecioBolsa(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim Palabra As String
    archivo = ArchivoInfTabEle(fecha)
    Dim i As Integer
    Dim j As Integer
    Dim Hora As Integer
    Dim Suma As Single
    i = 0
    
    ThisWorkbook.Worksheets("Precios Generaciones").Cells(1, 3).Value = fecha
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            textline = DejarUnSoloEspacio(textline)
            LArray = Split(textline, " ")
            If UBound(LArray) > -1 Then
                Palabra = Trim(UCase(LArray(0)))

                If UCase(Trim(LArray(0))) = "PRECIOS" Then
                    Line Input #1, textline
                    textline = DejarUnSoloEspacio(textline)
                    LArray = Split(textline, " ")
                    Palabra = Trim(UCase(LArray(0)))
                    i = 0
                    Suma = 0
                    Do While Palabra <> "PRECIO"
                        i = i + 1
                        PrecioBolsa(i) = LArray(1)
                        Suma = Suma + PrecioBolsa(i)
                        'Debug.Print PrecioBolsa(i)
                        ThisWorkbook.Worksheets("Precios Generaciones").Cells(i + 4, 2).Value = PrecioBolsa(i)
                        Line Input #1, textline
                        textline = DejarUnSoloEspacio(textline)
                        LArray = Split(textline, " ")
                        Palabra = Trim(UCase(LArray(0)))
                    Loop
                    ThisWorkbook.Worksheets("Precios Generaciones").Cells(i + 5, 2).Value = Suma / 24
                End If
            End If
        Loop
    Close #1
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerInfTabElePrecioBolsa"
End Sub

Sub LeerInfTabEleResultados(fecha As Date, Optional nroDiaAnterior As Integer)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim Palabra As String
    archivo = ArchivoInfTabEle(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Precios Generaciones").Cells(1, 9 - nroDiaAnterior).Value = fecha - nroDiaAnterior
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            textline = DejarUnSoloEspacio(textline)
            LArray = Split(textline, " ")
            If UBound(LArray) > -1 Then
                Palabra = Trim(UCase(LArray(0)))
                If UCase(Trim(LArray(0))) = "RESULTADOS" Then
                    Line Input #1, textline
                    textline = DejarUnSoloEspacio(textline)
                    LArray = Split(textline, "=")
                    Palabra = Trim(UCase(LArray(0)))
                    i = 0
                    Do While Palabra <> ""
                        i = i + 1
                        
                        Resultados(i).Resultado = LArray(0)
                        Resultados(i).Valor = LArray(1)
                        'Debug.Print Resultados(i).Resultado & "    " & Resultados(i).Valor
                        
                        ThisWorkbook.Worksheets("Precios Generaciones").Cells(i + 1, 9 - nroDiaAnterior).Value = Resultados(i).Valor
                        
                        Line Input #1, textline
                        textline = DejarUnSoloEspacio(textline)
                        LArray = Split(textline, "=")
                        Palabra = Trim(UCase(LArray(0)))
                    Loop
                End If
            End If
        Loop
    Close #1
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerInfTabEleResultados"
End Sub

Sub LeerInfTabEleDI(fecha As Date, Optional nroDiaAnterior As Integer)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim Palabra As String
    archivo = ArchivoInfTabEle(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Precios Generaciones").Cells(15, 9 - nroDiaAnterior).Value = fecha - nroDiaAnterior
    
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            textline = DejarUnSoloEspacio(textline)
            LArray = Split(textline, " ")
            If UBound(LArray) > -1 Then
                Palabra = Trim(UCase(LArray(0)))
                If UCase(Trim(LArray(0))) = "EMPRESAS" Then
                    Line Input #1, textline
                    textline = DejarUnSoloEspacio(textline)
                    LArray = Split(textline, " ")
                    Palabra = Trim(UCase(LArray(0)))
                    i = 0
                    Do While Palabra <> ""
                        i = i + 1
                        
                        Empresas(i).Empresa = LArray(0)
                        Empresas(i).DR = LArray(1)
                        Empresas(i).DI = LArray(2)
                        Empresas(i).MillPesos = LArray(3)
                        
                        ThisWorkbook.Worksheets("Precios Generaciones").Cells(15 + i, 9 - nroDiaAnterior).Value = Empresas(i).DI / 1000
                        
                        Line Input #1, textline
                        textline = DejarUnSoloEspacio(textline)
                        LArray = Split(textline, " ")
                        Palabra = Trim(UCase(LArray(0)))
                    Loop
                End If
            End If
        Loop
    Close #1
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerInfTabEleDI"
End Sub


Sub LeerInfTabEle(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim Palabra As String
    archivo = ArchivoInfTabEle(fecha)
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            textline = DejarUnSoloEspacio(textline)
            LArray = Split(textline, " ")
            If UBound(LArray) > -1 Then
                Palabra = Trim(UCase(LArray(0)))
                If UCase(Trim(LArray(0))) = "EMPRESAS" Then
                    Line Input #1, textline
                    textline = DejarUnSoloEspacio(textline)
                    LArray = Split(textline, " ")
                    Palabra = Trim(UCase(LArray(0)))
                    i = 0
                    Do While Palabra <> ""
                        i = i + 1
                        
                        Empresas(i).Empresa = LArray(0)
                        Empresas(i).DR = LArray(1)
                        Empresas(i).DI = LArray(2)
                        Empresas(i).MillPesos = LArray(3)
                        'Debug.Print Empresas(i).Empresa & "   "; Empresas(i).DR & "   "; Empresas(i).DI & "   "; Empresas(i).MillPesos
                        
                        Line Input #1, textline
                        textline = DejarUnSoloEspacio(textline)
                        LArray = Split(textline, " ")
                        Palabra = Trim(UCase(LArray(0)))
                    Loop
                ElseIf UCase(Trim(LArray(0))) = "PRECIOS" Then
                    Line Input #1, textline
                    textline = DejarUnSoloEspacio(textline)
                    LArray = Split(textline, " ")
                    Palabra = Trim(UCase(LArray(0)))
                    i = 0
                    Do While Palabra <> "PRECIO"
                        i = i + 1
                        
                        PrecioBolsa(i) = LArray(1)
                        'Debug.Print PrecioBolsa(i)
                        
                        Line Input #1, textline
                        textline = DejarUnSoloEspacio(textline)
                        LArray = Split(textline, " ")
                        Palabra = Trim(UCase(LArray(0)))
                    Loop
                ElseIf UCase(Trim(LArray(0))) = "RESULTADOS" Then
                    Line Input #1, textline
                    textline = DejarUnSoloEspacio(textline)
                    LArray = Split(textline, "=")
                    Palabra = Trim(UCase(LArray(0)))
                    i = 0
                    Do While Palabra <> ""
                        i = i + 1
                        
                        Resultados(i).Resultado = LArray(0)
                        Resultados(i).Valor = LArray(1)
                        'Debug.Print Resultados(i).Resultado & "    " & Resultados(i).Valor
                        
                        Line Input #1, textline
                        textline = DejarUnSoloEspacio(textline)
                        LArray = Split(textline, "=")
                        Palabra = Trim(UCase(LArray(0)))
                    Loop
                ElseIf UCase(Trim(LArray(0))) = "PRIORIDADES" Then
                    Line Input #1, textline
                    textline = DejarUnSoloEspacio(textline)
                    LArray = Split(textline, ";")
                    Palabra = Trim(UCase(LArray(0)))
                    i = 0
                    Do While Palabra <> ""
                        i = i + 1
                        
                        Prioridades(i).central = LArray(0)
                        Prioridades(i).Precio = LArray(1)
                        'Debug.Print Prioridades(i).Central & "    " & Prioridades(i).Precio
                        
                        Line Input #1, textline
                        textline = DejarUnSoloEspacio(textline)
                        LArray = Split(textline, ";")
                        Palabra = Trim(UCase(LArray(0)))
                    Loop
                End If
            End If
        Loop
    Close #1
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerInfTabEle"
End Sub