Option Explicit

Public FilaIDOGENERACION As Integer
Public FilaIDOAutoGenTermica As Integer
Public FilaIDOAutoGenHidraulica As Integer
Public FilaIDOGenCogenerador As Integer
Public FilaIDOGenTermica As Integer
Public FilaIDOGenHidraulica As Integer
Public FilaIDOGenMenHidraulica As Integer
Public FilaIDOGenMenTermica As Integer
Public FilaIDOGenMenEolica As Integer
Public FilaIDOGenMenSolar As Integer
Public FilaIDOImportaciones As Integer
Public FilaIDOExportaciones As Integer
Public FilaIDODispHidraulica As Integer
Public FilaIDODispTermica As Integer
Public FilaIDOAportes As Integer
Public FilaIDORio As Integer
Public FilaIDOTotalSINRio As Integer
Public FilaIDOReservas As Integer
Public FilaIDOEmbalse As Integer
Public FilaIDOTotalSINEmbalse As Integer
Public FilaIDOVertimientos As Integer
Public FilaIDOVertimientoEmbalse As Integer
Public FilaIDOTotalVertimiento As Integer

Public Const FilaParamGENERACION = 18
Public Const FilaParamAutoGenTermica = 19
Public Const FilaParamAutoGenHidraulica = 20
Public Const FilaParamGenCogenerador = 21
Public Const FilaParamGenTermica = 22
Public Const FilaParamGenHidraulica = 23
Public Const FilaParamGenMenHidraulica = 24
Public Const FilaParamGenMenTermica = 25
Public Const FilaParamGenMenEolica = 26
Public Const FilaParamGenMenSolar = 27
Public Const FilaParamImportaciones = 28
Public Const FilaParamExportaciones = 29
Public Const FilaParamDispHidraulica = 30
Public Const FilaParamDispTermica = 31
Public Const FilaParamAportes = 32
Public Const FilaParamRio = 33
Public Const FilaParamTotalSINRio = 34
Public Const FilaParamEmbalse = 35
Public Const FilaParamTotalSINEmbalse = 36
Public Const FilaParamVertimientoEmbalse = 37
Public Const FilaParamTotalVertimiento = 38


Public Const colIDOTexto = 1
Public Const colIDORedespacho = 2
Public Const colIDODespacho = 3
Public Const colIDOReal = 4
Public Const colIDOEmbVolUtilDiario = 4
Public Const colIDORioCaudalm3s = 2
Public Const colIDORioCaudalGWh = 3
Public Const colIDORioCaudalPorc = 4

Public Type typeGenIDO
    central As String
    RedespachoProgGWh As Single
    DespachoProgGWh As Single
    RealGWh As Single
End Type

Public Type typeEmbalseIDO
    embalse As String
    VolUtilDiario As Single
End Type

Public Type typeRioIDO
    Rio As String
    Caudalm3s As Single
    CaudalGWh As Single
    CaudalPorc As Single
End Type

Public Type typeIDOPalabraClave
    Codigo As String
    Palabra As String
    fila As Integer
End Type

Public GenIDO(MAXCENTRALES) As typeGenIDO
Public EmbalsesIDO(MAXEMBALSES) As typeEmbalseIDO
Public RiosIDO(MAXRIOS) As typeRioIDO
Public IDOPalabrasClave(50) As typeIDOPalabraClave

Sub LeerIDOGenProg(fecha As Date, Optional nroDiaAnterior As Integer)
    Dim prmLibro As String
    Dim FilaIDO As Integer
    Dim TextoIDO As String
    Dim regGenIDO As Integer
    Dim nmHoja As String

    
    prmLibro = ArchivoIDO(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Programado_Real").Cells(2, 2).Value = fecha - nroDiaAnterior
    
    Dim LibroIDO As Workbook
    Dim HojaIDO As Worksheet
    
    On Error GoTo ManejadorError
    Set LibroIDO = Workbooks.Open(prmLibro, ReadOnly:=True)
    'If LibroIDO.Sheets.Count > 1 Then nmHoja = NombreHojaIDO(fecha, nroDiaAnterior)
    nmHoja = LibroIDO.Sheets.Item(1).Name
    
    Set HojaIDO = LibroIDO.Worksheets(nmHoja)
    
    Dim fechaIDO As Date
    Dim strFechaIDO As String
    strFechaIDO = HojaIDO.Cells(3, 1).Value
    strFechaIDO = Mid(strFechaIDO, 7)
    fechaIDO = CDate(Replace(strFechaIDO, " de ", "/", 2))
    If fechaIDO <> fecha - nroDiaAnterior Then
        LogOfertaEPM "LeerIDOGEnProg: La fecha en el libro IDO " & prmLibro & " " & fechaIDO & Chr(10) & " no corresponde con la fecha del dia " & fecha - nroDiaAnterior
        LibroIDO.Close SaveChanges:=False
        Exit Sub
    End If
    
    LeerFilasPosicionPalabrasIDO HojaIDO
    regGenIDO = 0
    For FilaIDO = FilaIDOGENERACION To FilaIDOImportaciones
        TextoIDO = UCase(Trim(HojaIDO.Cells(FilaIDO, colIDOTexto).Value))
        If TextoIDO <> "" Then
            regGenIDO = regGenIDO + 1
            GenIDO(regGenIDO).central = TextoIDO
            GenIDO(regGenIDO).RedespachoProgGWh = HojaIDO.Cells(FilaIDO, colIDORedespacho).Value
            GenIDO(regGenIDO).DespachoProgGWh = HojaIDO.Cells(FilaIDO, colIDODespacho).Value
            GenIDO(regGenIDO).RealGWh = HojaIDO.Cells(FilaIDO, colIDOReal).Value
        End If
    Next FilaIDO
    
    LibroIDO.Close SaveChanges:=False
    

    Dim FilaProgContr As Integer
    Dim posCentralIDO As Integer
    Dim CentralIDO As String
    Dim CentralGenReal As String
    Dim regEquiv As Integer 'Indicador de posicion en el arreglo de equivalencias
    Dim Valor As Single
    
    OrdenarGenIDO GenIDO
    LeerEquivalencias
    
    FilaProgContr = 4
    CentralGenReal = ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 1).Value
    Do While CentralGenReal <> ""
        'Debug.Print CentralGenReal
        '---
        DoEvents
        regEquiv = 1
        Valor = 0
        
        If CentralGenReal <> "TOTAL TERMICAS" Then
            Do While regEquiv <= nroEquiv
                DoEvents
                If Equivalencias(regEquiv).informeGenProg = CentralGenReal Then
                    
                    CentralIDO = Equivalencias(regEquiv).CentralIDO
                    posCentralIDO = HallarPosCentralEnIDO(CentralIDO, GenIDO, regGenIDO)
                    If posCentralIDO <> -1 Then
                        Valor = Valor + GenIDO(posCentralIDO).RealGWh
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
                    
                    CentralIDO = Equivalencias(regEquiv).CentralIDO
                    posCentralIDO = HallarPosCentralEnIDO(CentralIDO, GenIDO, regGenIDO)
                    If posCentralIDO <> -1 Then
                        Valor = Valor + GenIDO(posCentralIDO).RealGWh
                        'Debug.Print CentralGenProg & " " & Equivalencias(regEquiv).CentralDDEC & " " & CStr(Valor)
                    Else
                    '    ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 3 - nroDiaAnterior).Value = "No hallada"
                    End If
    
                End If
            
                regEquiv = regEquiv + 1
            Loop
        End If
        
        ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 2).Value = Valor
        '---
        FilaProgContr = FilaProgContr + 1
        CentralGenReal = ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 1).Value
    Loop
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & prmLibro & " LeerIDOGenProg"
    
End Sub

Sub OrdenarGenIDO(prmGenIDO() As typeGenIDO)
    Dim nroCentral As Integer
    Dim nmCentral As String
   
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    nroCentral = 1
    nmCentral = prmGenIDO(nroCentral).central
    Do While nmCentral <> ""
        nroCentral = nroCentral + 1
        nmCentral = prmGenIDO(nroCentral).central
    Loop
    nroCentral = nroCentral - 1
    
    'Ordenar centrales por nombre
    For i = 1 To nroCentral - 1
        For j = i + 1 To nroCentral
            If prmGenIDO(i).central > prmGenIDO(j).central Then
            
                prmGenIDO(0).central = prmGenIDO(i).central
                prmGenIDO(0).RedespachoProgGWh = prmGenIDO(i).RedespachoProgGWh
                prmGenIDO(0).DespachoProgGWh = prmGenIDO(i).DespachoProgGWh
                prmGenIDO(0).RealGWh = prmGenIDO(i).RealGWh
                
                prmGenIDO(i).central = prmGenIDO(j).central
                prmGenIDO(i).RedespachoProgGWh = prmGenIDO(j).RedespachoProgGWh
                prmGenIDO(i).DespachoProgGWh = prmGenIDO(j).DespachoProgGWh
                prmGenIDO(i).RealGWh = prmGenIDO(j).RealGWh
                
                prmGenIDO(j).central = prmGenIDO(0).central
                prmGenIDO(j).RedespachoProgGWh = prmGenIDO(0).RedespachoProgGWh
                prmGenIDO(j).DespachoProgGWh = prmGenIDO(0).DespachoProgGWh
                prmGenIDO(j).RealGWh = prmGenIDO(0).RealGWh

            End If
        Next j
    Next i

End Sub

Sub OrdenarEmbalsesIDO(prmEmbIDO() As typeEmbalseIDO)
    Dim nroEmbalse As Integer
    Dim nmEmbalse As String
   
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    nroEmbalse = 1
    nmEmbalse = prmEmbIDO(nroEmbalse).embalse
    Do While nmEmbalse <> ""
        nroEmbalse = nroEmbalse + 1
        nmEmbalse = prmEmbIDO(nroEmbalse).embalse
    Loop
    nroEmbalse = nroEmbalse - 1
    
    'Ordenar centrales por nombre
    For i = 1 To nroEmbalse - 1
        For j = i + 1 To nroEmbalse
            If prmEmbIDO(i).embalse > prmEmbIDO(j).embalse Then
            
                prmEmbIDO(0).embalse = prmEmbIDO(i).embalse
                prmEmbIDO(0).VolUtilDiario = prmEmbIDO(i).VolUtilDiario

                prmEmbIDO(i).embalse = prmEmbIDO(j).embalse
                prmEmbIDO(i).VolUtilDiario = prmEmbIDO(j).VolUtilDiario

                prmEmbIDO(j).embalse = prmEmbIDO(0).embalse
                prmEmbIDO(j).VolUtilDiario = prmEmbIDO(0).VolUtilDiario

            End If
        Next j
    Next i

End Sub

Sub LeerIDOGenRealEmpresa(fecha As Date, Optional nroDiaAnterior As Integer)
    Dim archivo As String

    Dim prmLibro As String
    Dim FilaIDO As Integer
    Dim TextoIDO As String
    Dim regGenIDO As Integer
    Dim nmHoja As String
    Dim LibroIDO As Workbook
    Dim HojaIDO As Worksheet
    
    prmLibro = ArchivoIDO(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Precios Generaciones").Cells(24, 9 - nroDiaAnterior).Value = fecha - nroDiaAnterior
    
    
    On Error GoTo ManejadorError
    Set LibroIDO = Workbooks.Open(prmLibro, ReadOnly:=True)
    'nmHoja = NombreHojaIDO(fecha, nroDiaAnterior)
    nmHoja = LibroIDO.Sheets.Item(1).Name
    Set HojaIDO = LibroIDO.Worksheets(nmHoja)
    
    Dim fechaIDO As Date
    Dim strFechaIDO As String
    strFechaIDO = HojaIDO.Cells(3, 1).Value
    strFechaIDO = Mid(strFechaIDO, 7)
    fechaIDO = CDate(Replace(strFechaIDO, " de ", "/", 2))
    If fechaIDO <> fecha - nroDiaAnterior Then
        LogOfertaEPM "LeerIDOGenRealEmpresa: La fecha en el libro IDO " & prmLibro & " " & fechaIDO & Chr(10) & " no corresponde con la fecha del dia " & fecha - nroDiaAnterior
        LibroIDO.Close SaveChanges:=False
        Exit Sub
    End If
    
    
    LeerFilasPosicionPalabrasIDO HojaIDO
    regGenIDO = 0
    For FilaIDO = FilaIDOGENERACION To FilaIDOImportaciones
        TextoIDO = UCase(Trim(HojaIDO.Cells(FilaIDO, colIDOTexto).Value))
        If TextoIDO <> "" Then
            regGenIDO = regGenIDO + 1
            GenIDO(regGenIDO).central = TextoIDO
            GenIDO(regGenIDO).RedespachoProgGWh = HojaIDO.Cells(FilaIDO, colIDORedespacho).Value
            GenIDO(regGenIDO).DespachoProgGWh = HojaIDO.Cells(FilaIDO, colIDODespacho).Value
            GenIDO(regGenIDO).RealGWh = HojaIDO.Cells(FilaIDO, colIDOReal).Value
        End If
    Next FilaIDO
    
    LibroIDO.Close SaveChanges:=False
    
    Dim FilaPreciosGen As Integer
    Dim posCentralIDO As Integer
    Dim CentralIDO As String
    Dim CentralGenReal As String
    Dim regEquiv As Integer 'Indicador de posicion en el arreglo de equivalencias
    Dim Valor As Single
    
    OrdenarGenIDO GenIDO
    LeerEquivalencias
    
    FilaPreciosGen = 25
    CentralGenReal = ThisWorkbook.Worksheets("Precios Generaciones").Cells(FilaPreciosGen, 5).Value
    Do While CentralGenReal <> ""
        'Debug.Print CentralGenReal
        '---
        DoEvents
        regEquiv = 1
        Valor = 0
        

            Do While regEquiv <= nroEquiv
                DoEvents
                If Equivalencias(regEquiv).informeGenRealEmp = CentralGenReal Then
                    
                    CentralIDO = Equivalencias(regEquiv).CentralIDO
                    posCentralIDO = HallarPosCentralEnIDO(CentralIDO, GenIDO, regGenIDO)
                    If posCentralIDO <> -1 Then
                        Valor = Valor + GenIDO(posCentralIDO).RealGWh
                        'Debug.Print CentralGenReal & " " & Equivalencias(regEquiv).CentralIDO & " " & CStr(Valor)
                    Else
                    '    ThisWorkbook.Worksheets("Programado_Real").Cells(FilaProgContr, 3 - nroDiaAnterior).Value = "No hallada"
                    End If
    
                End If
            
                regEquiv = regEquiv + 1
            Loop

        
        ThisWorkbook.Worksheets("Precios Generaciones").Cells(FilaPreciosGen, 9 - nroDiaAnterior).Value = Valor
        '---
        FilaPreciosGen = FilaPreciosGen + 1
        CentralGenReal = ThisWorkbook.Worksheets("Precios Generaciones").Cells(FilaPreciosGen, 5).Value
    Loop
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerIDOGenRealEmpresa"
End Sub

Public Function HallarPosCentralEnIDO(nmCentral As String, prmGenIDO() As typeGenIDO, nroCentrales As Integer) As Integer
    Dim centralInf As String
    Dim centralSup As String
    Dim centralMed As String
    Dim blnHallado As String
    Dim intSup As Integer
    Dim intInf As Integer
    Dim intMed As Integer
    
    nmCentral = UCase(Trim(nmCentral))
    
    HallarPosCentralEnIDO = -1
    
    intSup = nroCentrales
    intInf = 1
    intMed = Int((intSup + intInf) / 2)
    centralInf = UCase(Trim(prmGenIDO(intInf).central))
    centralSup = UCase(Trim(prmGenIDO(intSup).central))
    centralMed = UCase(Trim(prmGenIDO(intMed).central))
    
    If centralInf = nmCentral Then
        HallarPosCentralEnIDO = intInf
        Exit Function
    End If
    
    If centralSup = nmCentral Then
        HallarPosCentralEnIDO = intSup
        Exit Function
    End If
    
    If centralMed = nmCentral Then
        HallarPosCentralEnIDO = intMed
        Exit Function
    End If
    
    blnHallado = False
    
    Do While centralMed <> nmCentral And blnHallado = False And (intSup - intInf > 1)
        
        DoEvents
        
        If nmCentral > centralMed Then
            intInf = intMed
            centralInf = UCase(Trim(prmGenIDO(intInf).central))
        End If
        
        If nmCentral < centralMed Then
            intSup = intMed
            centralSup = UCase(Trim(prmGenIDO(intSup).central))
        End If
        
        intMed = Int((intSup + intInf) / 2)
        centralMed = UCase(Trim(prmGenIDO(intMed).central))
        
        If nmCentral = centralMed Then
            HallarPosCentralEnIDO = intMed
            blnHallado = True
        End If
    Loop

End Function


Sub LeerFilasPosicionPalabrasIDO(HojaIDO As Worksheet)
    Dim codPalabra As String
    Dim Palabra As String
    Dim PalabraIDO As String
    Dim FilaIDO As Integer
    Dim FilaParam As Integer
    Dim FilaIDOActual As Integer
    Dim blnHallado As Boolean
    Dim ultimaFilaHallada As Integer
    
    FilaIDOActual = 1
    
    'hojaIDO.UsedRange.MergeCells = False
    
    FilaParam = FilaParamIDOPalabras
    codPalabra = ThisWorkbook.Worksheets("Parametros").Cells(FilaParam, ColParamCodPalabra).Value
    Palabra = ThisWorkbook.Worksheets("Parametros").Cells(FilaParam, ColParamPalabraIDO).Value
    Do While Trim(codPalabra) <> ""

        
        blnHallado = False
        PalabraIDO = HojaIDO.Cells(FilaIDOActual, 1).Value
        Do While (blnHallado = False And FilaIDOActual < 1000)
            
            PalabraIDO = Replace(PalabraIDO, Chr(160), Chr(32)) 'Hay un espacio en blanco que corresponde al caracter Chr(160)
            PalabraIDO = Trim(UCase(PalabraIDO))
            If PalabraIDO = UCase(Trim(Palabra)) Then
                blnHallado = True
                ultimaFilaHallada = FilaIDOActual
                ThisWorkbook.Worksheets("Parametros").Cells(FilaParam, ColParamFilaIDO).Value = FilaIDOActual
                'Debug.Print codPalabra & "  " & CStr(FilaIDOActual) & "  " & Palabra
            End If
            
            FilaIDOActual = FilaIDOActual + 1
            PalabraIDO = HojaIDO.Cells(FilaIDOActual, 1).Value
        Loop
        
        If blnHallado = False Then
            FilaIDOActual = ultimaFilaHallada + 1
            ThisWorkbook.Worksheets("Parametros").Cells(FilaParam, ColParamFilaIDO).Value = -1
        End If

        FilaParam = FilaParam + 1
        codPalabra = ThisWorkbook.Worksheets("Parametros").Cells(FilaParam, ColParamCodPalabra).Value
        Palabra = ThisWorkbook.Worksheets("Parametros").Cells(FilaParam, ColParamPalabraIDO).Value
    Loop
    FilaIDOGENERACION = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamGENERACION, ColParamFilaIDO).Value
    FilaIDOAutoGenTermica = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamAutoGenTermica, ColParamFilaIDO).Value
    FilaIDOAutoGenHidraulica = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamAutoGenHidraulica, ColParamFilaIDO).Value
    FilaIDOGenCogenerador = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamGenCogenerador, ColParamFilaIDO).Value
    FilaIDOGenTermica = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamGenTermica, ColParamFilaIDO).Value
    FilaIDOGenHidraulica = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamGenHidraulica, ColParamFilaIDO).Value
    FilaIDOGenMenHidraulica = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamGenMenHidraulica, ColParamFilaIDO).Value
    FilaIDOGenMenTermica = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamGenMenTermica, ColParamFilaIDO).Value
    FilaIDOGenMenEolica = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamGenMenEolica, ColParamFilaIDO).Value
    FilaIDOGenMenSolar = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamGenMenSolar, ColParamFilaIDO).Value
    FilaIDOImportaciones = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamImportaciones, ColParamFilaIDO).Value
    FilaIDOExportaciones = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamExportaciones, ColParamFilaIDO).Value
    FilaIDODispHidraulica = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamDispHidraulica, ColParamFilaIDO).Value
    FilaIDODispTermica = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamDispTermica, ColParamFilaIDO).Value
    FilaIDOAportes = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamAportes, ColParamFilaIDO).Value
    FilaIDORio = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRio, ColParamFilaIDO).Value
    FilaIDOTotalSINRio = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamTotalSINRio, ColParamFilaIDO).Value
    FilaIDOEmbalse = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamEmbalse, ColParamFilaIDO).Value
    FilaIDOTotalSINEmbalse = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamTotalSINEmbalse, ColParamFilaIDO).Value
    FilaIDOVertimientoEmbalse = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamVertimientoEmbalse, ColParamFilaIDO).Value
    FilaIDOTotalVertimiento = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamTotalVertimiento, ColParamFilaIDO).Value

End Sub

Public Function ArchivoIDO(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    Dim ext As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamIDO, ColParamPrefijo).Value
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoIDO = raiz & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha)
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamIDO, ColParamRaiz).Value
        ArchivoIDO = raiz & año & "\" & mes & "\" & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha)
    End If
    ext = ExtensionExcel(ArchivoIDO)
    ArchivoIDO = ArchivoIDO & ext
End Function


Public Function ArchivoIDOenDiario(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    Dim ext As String
    raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamDIARIO, ColParamRaiz).Value
    prefijo = "IDO"
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    ArchivoIDOenDiario = raiz & "\" & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha)
    ext = ExtensionExcel(ArchivoIDOenDiario)
    ArchivoIDOenDiario = ArchivoIDOenDiario & ext
End Function


Public Function NombreHojaIDO(fecha As Date, Optional nroDiaAnterior As Integer) As String
    Dim prefijo As String
    Dim mes As String
    Dim auxFecha As Date
    auxFecha = fecha - nroDiaAnterior
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamIDO, ColParamPrefijo).Value
    mes = NombreMes(nmMes.largo, auxFecha)
    NombreHojaIDO = prefijo & NombreMes(NumeroConCero, auxFecha) & nroDia(ConCero, auxFecha)
End Function

Sub LeerIDOEmbalses(fecha As Date, Optional nroDiaAnterior As Integer)
    Dim archivo As String

    Dim prmLibro As String
    Dim FilaIDO As Integer
    Dim TextoIDO As String
    Dim regEmbIDO As Integer
    Dim nmHoja As String
    Dim LibroIDO As Workbook
    Dim HojaIDO As Worksheet
    
    prmLibro = ArchivoIDO(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Embalses").Cells(1, ColEmbVolInicialPorc).Value = fecha - nroDiaAnterior
    
    On Error GoTo ManejadorError
    Set LibroIDO = Workbooks.Open(prmLibro, ReadOnly:=True)
    'nmHoja = NombreHojaIDO(fecha, nroDiaAnterior)
    nmHoja = LibroIDO.Sheets.Item(1).Name
    Set HojaIDO = LibroIDO.Worksheets(nmHoja)
    
    Dim fechaIDO As Date
    Dim strFechaIDO As String
    strFechaIDO = HojaIDO.Cells(3, 1).Value
    strFechaIDO = Mid(strFechaIDO, 7)
    fechaIDO = CDate(Replace(strFechaIDO, " de ", "/", 2))
    If fechaIDO <> fecha - nroDiaAnterior Then
        LogOfertaEPM "LeerIDOEmbalses: La fecha en el libro IDO " & prmLibro & " " & fechaIDO & Chr(10) & " no corresponde con la fecha del dia " & fecha - nroDiaAnterior
        LibroIDO.Close SaveChanges:=False
        Exit Sub
    End If
    
    
    LeerFilasPosicionPalabrasIDO HojaIDO
    regEmbIDO = 0
    For FilaIDO = FilaIDOEmbalse + 1 To FilaIDOTotalSINEmbalse
        TextoIDO = UCase(Trim(HojaIDO.Cells(FilaIDO, colIDOTexto).Value))
        If TextoIDO <> "" Then
            regEmbIDO = regEmbIDO + 1
            EmbalsesIDO(regEmbIDO).embalse = TextoIDO
            EmbalsesIDO(regEmbIDO).VolUtilDiario = HojaIDO.Cells(FilaIDO, colIDOEmbVolUtilDiario).Value
            'Debug.Print EmbalsesIDO(regEmbIDO).Embalse & "  " & EmbalsesIDO(regEmbIDO).VolUtilDiario
        End If
    Next FilaIDO
    
    LibroIDO.Close SaveChanges:=False
    
    Dim FilaEmbalse As Integer
    Dim posEmbalseIDO As Integer
    Dim EmbalseIDO As String
    Dim EmbalseBalance As String
    Dim Valor As Single

    OrdenarEmbalsesIDO EmbalsesIDO

    FilaEmbalse = 3
    EmbalseIDO = ThisWorkbook.Worksheets("Embalses").Cells(FilaEmbalse, ColEmbEmbalseIDO).Value
    Do While EmbalseIDO <> ""
        'Debug.Print EmbalseIDO

        DoEvents

        Valor = 0

        posEmbalseIDO = HallarPosEmbalseEnIDO(EmbalseIDO, EmbalsesIDO, regEmbIDO)
        If posEmbalseIDO <> -1 Then
            Valor = Valor + EmbalsesIDO(posEmbalseIDO).VolUtilDiario
            'Debug.Print EmbalseIDO & " " & CStr(Valor)
        Else
            ThisWorkbook.Worksheets("Embalses").Cells(FilaEmbalse, ColEmbVolInicialPorc).Value = "No hallada"
        End If


        ThisWorkbook.Worksheets("Embalses").Cells(FilaEmbalse, ColEmbVolInicialPorc).Value = Valor

        FilaEmbalse = FilaEmbalse + 1
        EmbalseIDO = ThisWorkbook.Worksheets("Embalses").Cells(FilaEmbalse, ColEmbEmbalseIDO).Value
    Loop
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerIDOEmbalses"
End Sub

Public Function HallarPosEmbalseEnIDO(nmEmbalse As String, prmEmbIDO() As typeEmbalseIDO, nroEmbalses As Integer) As Integer
    Dim EmbalseInf As String
    Dim EmbalseSup As String
    Dim EmbalseMed As String
    Dim blnHallado As String
    Dim intSup As Integer
    Dim intInf As Integer
    Dim intMed As Integer
    
    nmEmbalse = UCase(Trim(nmEmbalse))
    
    HallarPosEmbalseEnIDO = -1
    
    intSup = nroEmbalses
    intInf = 1
    intMed = Int((intSup + intInf) / 2)
    EmbalseInf = UCase(Trim(prmEmbIDO(intInf).embalse))
    EmbalseSup = UCase(Trim(prmEmbIDO(intSup).embalse))
    EmbalseMed = UCase(Trim(prmEmbIDO(intMed).embalse))
    
    If EmbalseInf = nmEmbalse Then
        HallarPosEmbalseEnIDO = intInf
        Exit Function
    End If
    
    If EmbalseSup = nmEmbalse Then
        HallarPosEmbalseEnIDO = intSup
        Exit Function
    End If
    
    If EmbalseMed = nmEmbalse Then
        HallarPosEmbalseEnIDO = intMed
        Exit Function
    End If
    
    blnHallado = False
    
    Do While EmbalseMed <> nmEmbalse And blnHallado = False And (intSup - intInf > 1)
        
        DoEvents
        
        If nmEmbalse > EmbalseMed Then
            intInf = intMed
            EmbalseInf = UCase(Trim(prmEmbIDO(intInf).embalse))
        End If
        
        If nmEmbalse < EmbalseMed Then
            intSup = intMed
            EmbalseSup = UCase(Trim(prmEmbIDO(intSup).embalse))
        End If
        
        intMed = Int((intSup + intInf) / 2)
        EmbalseMed = UCase(Trim(prmEmbIDO(intMed).embalse))
        
        If nmEmbalse = EmbalseMed Then
            HallarPosEmbalseEnIDO = intMed
            blnHallado = True
        End If
    Loop

End Function

Sub LeerIDORios(fecha As Date, Optional nroDiaAnterior As Integer)
    Dim archivo As String

    Dim prmLibro As String
    Dim FilaIDO As Integer
    Dim TextoIDO As String
    Dim regRioIDO As Integer
    Dim nmHoja As String
    Dim LibroIDO As Workbook
    Dim HojaIDO As Worksheet
    
    prmLibro = ArchivoIDO(fecha - nroDiaAnterior)
    ThisWorkbook.Worksheets("Rios").Cells(1, ColRiosCaudal).Value = fecha - nroDiaAnterior
    
    On Error GoTo ManejadorError
    Set LibroIDO = Workbooks.Open(prmLibro, ReadOnly:=True)
    'nmHoja = NombreHojaIDO(fecha, nroDiaAnterior)
    nmHoja = LibroIDO.Sheets.Item(1).Name
    Set HojaIDO = LibroIDO.Worksheets(nmHoja)
    
    Dim fechaIDO As Date
    Dim strFechaIDO As String
    strFechaIDO = HojaIDO.Cells(3, 1).Value
    strFechaIDO = Mid(strFechaIDO, 7)
    fechaIDO = CDate(Replace(strFechaIDO, " de ", "/", 2))
    If fechaIDO <> fecha - nroDiaAnterior Then
        LogOfertaEPM "LeerIDORios: La fecha en el libro IDO " & prmLibro & " " & fechaIDO & Chr(10) & " no corresponde con la fecha del dia " & fecha - nroDiaAnterior
        LibroIDO.Close SaveChanges:=False
        Exit Sub
    End If
    
    LeerFilasPosicionPalabrasIDO HojaIDO
    regRioIDO = 0
    For FilaIDO = FilaIDORio + 1 To FilaIDOEmbalse - 1
        TextoIDO = UCase(Trim(HojaIDO.Cells(FilaIDO, colIDOTexto).Value))
        If TextoIDO <> "" Then
            regRioIDO = regRioIDO + 1
            RiosIDO(regRioIDO).Rio = TextoIDO
            RiosIDO(regRioIDO).Caudalm3s = HojaIDO.Cells(FilaIDO, colIDORioCaudalm3s).Value
            RiosIDO(regRioIDO).CaudalGWh = HojaIDO.Cells(FilaIDO, colIDORioCaudalGWh).Value
            RiosIDO(regRioIDO).CaudalPorc = HojaIDO.Cells(FilaIDO, colIDORioCaudalPorc).Value
            'Debug.Print RiosIDO(regRioIDO).Rio & "  " & RiosIDO(regRioIDO).Caudalm3s & " PORC = " & RiosIDO(regRioIDO).CaudalPorc
        End If
    Next FilaIDO
    
    LibroIDO.Close SaveChanges:=False
    
    Dim FilaRios As Integer
    Dim posRioIDO As Integer
    Dim RioIDO As String

    Dim Valor As Single
    Dim porc As Single
    Dim GWhDia As Single

    OrdenarRiosIDO RiosIDO

    FilaRios = 3
    RioIDO = ThisWorkbook.Worksheets("Rios").Cells(FilaRios, ColRiosRioIDO).Value
    Do While RioIDO <> ""
        'Debug.Print RioIDO

        DoEvents

        Valor = 0
        porc = 0
        posRioIDO = HallarPosRioEnIDO(RioIDO, RiosIDO, regRioIDO)
        If posRioIDO <> -1 Then
            Valor = Valor + RiosIDO(posRioIDO).Caudalm3s
            porc = RiosIDO(posRioIDO).CaudalPorc
            GWhDia = RiosIDO(posRioIDO).CaudalGWh
            'Debug.Print RioIDO & " " & CStr(valor) & " Porc " & CStr(porc)
        Else
            ThisWorkbook.Worksheets("Rios").Cells(FilaRios, ColRiosCaudal).Value = "No hallada"
        End If


        ThisWorkbook.Worksheets("Rios").Cells(FilaRios, ColRiosCaudal).Value = Valor
        ThisWorkbook.Worksheets("Rios").Cells(FilaRios, ColRiosPorc).Value = porc
        ThisWorkbook.Worksheets("Rios").Cells(FilaRios, ColRiosGWhDia).Value = GWhDia

        FilaRios = FilaRios + 1
        RioIDO = ThisWorkbook.Worksheets("Rios").Cells(FilaRios, ColRiosRioIDO).Value
    Loop
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerIDORios"
End Sub

Sub OrdenarRiosIDO(prmRiosIDO() As typeRioIDO)
    Dim nroRio As Integer
    Dim nmRio As String
   
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    nroRio = 1
    nmRio = prmRiosIDO(nroRio).Rio
'    Do While nmRio <> ""
'        nroRio = nroRio + 1
'        nmRio = prmRiosIDO(nroRio).Rio
'    Loop
    For i = 1 To MAXRIOS
        nmRio = prmRiosIDO(nroRio).Rio
        If Trim(nmRio) <> "" Then
            nroRio = i
        End If
    Next i
    nroRio = nroRio - 1
    
    'Ordenar centrales por nombre
    For i = 1 To nroRio - 1
        For j = i + 1 To nroRio
            If prmRiosIDO(i).Rio > prmRiosIDO(j).Rio Then
            
                prmRiosIDO(0).Rio = prmRiosIDO(i).Rio
                prmRiosIDO(0).Caudalm3s = prmRiosIDO(i).Caudalm3s
                prmRiosIDO(0).CaudalGWh = prmRiosIDO(i).CaudalGWh
                prmRiosIDO(0).CaudalPorc = prmRiosIDO(i).CaudalPorc

                prmRiosIDO(i).Rio = prmRiosIDO(j).Rio
                prmRiosIDO(i).Caudalm3s = prmRiosIDO(j).Caudalm3s
                prmRiosIDO(i).CaudalGWh = prmRiosIDO(j).CaudalGWh
                prmRiosIDO(i).CaudalPorc = prmRiosIDO(j).CaudalPorc

                prmRiosIDO(j).Rio = prmRiosIDO(0).Rio
                prmRiosIDO(j).Caudalm3s = prmRiosIDO(0).Caudalm3s
                prmRiosIDO(j).CaudalGWh = prmRiosIDO(0).CaudalGWh
                prmRiosIDO(j).CaudalPorc = prmRiosIDO(0).CaudalPorc
                

            End If
        Next j
    Next i

End Sub

Public Function HallarPosRioEnIDO(nmRio As String, prmRioIDO() As typeRioIDO, nroRios As Integer) As Integer
    Dim RioInf As String
    Dim RioSup As String
    Dim RioMed As String
    Dim blnHallado As String
    Dim intSup As Integer
    Dim intInf As Integer
    Dim intMed As Integer
    
    nmRio = UCase(Trim(nmRio))
    
    HallarPosRioEnIDO = -1
    
    intSup = nroRios
    intInf = 1
    intMed = Int((intSup + intInf) / 2)
    RioInf = UCase(Trim(prmRioIDO(intInf).Rio))
    RioSup = UCase(Trim(prmRioIDO(intSup).Rio))
    RioMed = UCase(Trim(prmRioIDO(intMed).Rio))
    
    If RioInf = nmRio Then
        HallarPosRioEnIDO = intInf
        Exit Function
    End If
    
    If RioSup = nmRio Then
        HallarPosRioEnIDO = intSup
        Exit Function
    End If
    
    If RioMed = nmRio Then
        HallarPosRioEnIDO = intMed
        Exit Function
    End If
    
    blnHallado = False
    
    Do While RioMed <> nmRio And blnHallado = False And (intSup - intInf > 1)
        
        DoEvents
        
        If nmRio > RioMed Then
            intInf = intMed
            RioInf = UCase(Trim(prmRioIDO(intInf).Rio))
        End If
        
        If nmRio < RioMed Then
            intSup = intMed
            RioSup = UCase(Trim(prmRioIDO(intSup).Rio))
        End If
        
        intMed = Int((intSup + intInf) / 2)
        RioMed = UCase(Trim(prmRioIDO(intMed).Rio))
        
        If nmRio = RioMed Then
            HallarPosRioEnIDO = intMed
            blnHallado = True
        End If
    Loop

End Function

Public Sub CopiarArchivoIDOdesdeDiario(fecha As Date)
    Dim año As String
    Dim mes As String
    Dim raiz As String
    Dim fso As Object
    'Dim CarpetaDestino As Object
    Dim RutaDestino As String
    
    Dim strArchivoIDO As String
    Dim strArchivoIDOenDiario As String
    
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        RutaDestino = raiz
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamIDO, ColParamRaiz).Value
        RutaDestino = raiz & año & "\" & mes & "\"
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
  
    strArchivoIDO = ArchivoIDO(fecha)
    
    If Not (Len(Dir(strArchivoIDO)) > 0) Then 'Si no esta el archivo lo copia desde la carpeta Diario si esta.
        strArchivoIDOenDiario = ArchivoIDOenDiario(fecha)
        If Len(Dir(strArchivoIDOenDiario)) > 0 Then
            fso.copyfile strArchivoIDOenDiario, RutaDestino
        Else
            LogOfertaEPM "Falta " & strArchivoIDOenDiario
        End If
    End If
    
    Set fso = Nothing
End Sub
