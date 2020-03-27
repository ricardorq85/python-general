Option Explicit

Public nroUnidades As Integer

Public Type typeOFEIDispUni
    unidad As String
    Planta As String
    MWh(24) As Single
End Type

Public Type typeOFEIDispCen
    Planta As String
    MWh(24) As Single
    MO(24) As Single
    AGC(24) As Single
    MX(24) As Single
    MW(24) As Single
    Total As Single
End Type


Public Type TypeOFEIAGCU
    unidad As String
    Planta As String
    MWh(24) As Single
End Type

Public Type TypeOFEIAGCP
    Planta As String
    MWh(24) As Single
End Type

Public Type TypeOFEICombP
    Planta As String
    comb As String
End Type

Public Type TypeOFEIPrueba  'Las plantas o sus unidades pueden estar en prueba
    recurso As String
    prueba(24) As Integer
End Type

Public Type TypeOFEIConfigP
    Planta As String
    config As String
End Type

Public Type TypeOFEIMOP
    Planta As String
    MWh(24) As Single
End Type


Public DispUni(MAXUNIDADES) As typeOFEIDispUni
Public DispCen(30000) As typeOFEIDispCen   'Se localiza el registro con el codigoHash de la central
Public DispAGCU(MAXUNIDADES) As TypeOFEIAGCU
Public DispAGCP(MAXCENTRALES) As TypeOFEIAGCP
Public CombPlanta(MAXCENTRALES) As TypeOFEICombP
Public PruebasPlanta(MAXCENTRALES) As TypeOFEIPrueba
Public ConfigPlanta(MAXCENTRALES) As TypeOFEIConfigP
Public MinOpePlanta(MAXCENTRALES) As TypeOFEIMOP


Sub LeerOfeiDispUnidad(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoOFEI(fecha)
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 25 Then
                If Trim(UCase(LArray(1))) = "D" Then
                    Debug.Print textline
                    DispUni(i).unidad = Trim(LArray(0))
                    For Hora = 1 To 24
                        DispUni(i).MWh(Hora) = LArray(Hora + 1)
                    Next Hora
                    i = i + 1
                End If
            End If
        Loop
    Close #1
End Sub
'Escribe la disponibilidad de la planta en la hoja Generacion
Sub LeerOfeiDispPlanta(fecha As Date, Optional Escribir As Boolean = True)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    
    Dim nmCentral As String
    
    archivo = ArchivoOFEI(fecha)
    Dim i As Integer
    Dim Hora As Integer
    Dim codHashCen As Integer

    ThisWorkbook.Worksheets("Generacion").Cells(1, 5).Value = fecha
    
    nroUnidades = LeerCentralUnidad

    For i = 0 To 30000
        DispCen(i).Total = 0
        For Hora = 1 To 24
            DispCen(i).MWh(Hora) = 0
            DispCen(i).AGC(Hora) = 0
            DispCen(i).MX(Hora) = 0
            DispCen(i).MW(Hora) = 0
        Next Hora
    Next i

    i = 0
    
    On Error GoTo ManejadorError
    
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 25 Then
                If Trim(UCase(LArray(1))) = "D" Then
                    
                    'Debug.Print textline
                    DispUni(i).unidad = Trim(LArray(0))
                    nmCentral = UCase(Trim(nmCentralUnidad(DispUni(i).unidad, CentralDeUnidad, nroUnidades)))
                    codHashCen = CodigoHash(nmCentral)
                    DispUni(i).Planta = nmCentral
                    DispCen(codHashCen).Planta = nmCentral
                    For Hora = 1 To 24
                        DispUni(i).MWh(Hora) = LArray(Hora + 1)
                        DispCen(codHashCen).MWh(Hora) = DispCen(codHashCen).MWh(Hora) + DispUni(i).MWh(Hora)
                        DispCen(codHashCen).Total = DispCen(codHashCen).Total + DispUni(i).MWh(Hora)
                    Next Hora
                    
                    i = i + 1
                    
                End If
                If Trim(UCase(LArray(1))) = "MO" Then
                    nmCentral = CenOFEIaDDEC(Trim(LArray(0)))
                    codHashCen = CodigoHash(nmCentral)
                    For Hora = 1 To 24
                        DispCen(codHashCen).MO(Hora) = LArray(Hora + 1)
                    Next Hora
                    'Debug.Print codHashCen; "  "; textline
                    i = i + 1

                End If
            End If
        Loop
    Close #1
    
    If Escribir Then
        Dim FilaGen As Integer
        FilaGen = 3
        nmCentral = UCase(Trim(ThisWorkbook.Worksheets("Generacion").Cells(FilaGen, 1).Value))
        Do While nmCentral <> ""
            DoEvents
                codHashCen = CodigoHash(nmCentral)
                ThisWorkbook.Worksheets("Generacion").Cells(FilaGen, 5).Value = DispCen(codHashCen).Total / 24
            FilaGen = FilaGen + 1
            nmCentral = UCase(Trim(ThisWorkbook.Worksheets("Generacion").Cells(FilaGen, 1).Value))
        Loop
    End If
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerOfeiDispPlanta"
End Sub



Sub LeerOfeiAgcUnidad(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoOFEI(fecha)
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 25 Then
                If Trim(UCase(LArray(1))) = "AGCU" Then
                    'Debug.Print textline
                    DispAGCU(i).unidad = Trim(LArray(0))
                    For Hora = 1 To 24
                        DispAGCU(i).MWh(Hora) = LArray(Hora + 1)
                    Next Hora
                    i = i + 1
                End If
            End If
        Loop
    Close #1
End Sub

Sub LeerOfeiAgcPlanta(fecha As Date, Optional posX As Integer = 0, Optional posY As Integer = 0)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoOFEI(fecha)
    Dim nroCenAGC As Integer
    Dim j As Integer
    Dim Hora As Integer
    nroCenAGC = 0
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 25 Then
                If Trim(UCase(LArray(1))) = "AGCP" Then
                    'Debug.Print textline
                    nroCenAGC = nroCenAGC + 1
                    DispAGCP(nroCenAGC).Planta = Trim(LArray(0))
                    For Hora = 1 To 24
                        DispAGCP(nroCenAGC).MWh(Hora) = LArray(Hora + 1)
                    Next Hora
                    
                End If
            End If
        Loop
    Close #1
    
    Dim TotalCentral As Double
    
    

    ThisWorkbook.Worksheets("Servicio AGC").Cells(1 + posX, 1 + posY).Value = fecha
    ThisWorkbook.Worksheets("Servicio AGC").Cells(2 + posX, 1 + posY).Value = "Hora"
    ThisWorkbook.Worksheets("Servicio AGC").Cells(1 + posX, 1 + posY).Value = fecha
    ThisWorkbook.Worksheets("Servicio AGC").Cells(1 + posX, 2 + posY).Value = "Oferta de AGC por planta(MWh)"
    ThisWorkbook.Worksheets("Servicio AGC").Range(Cells(1 + posX, 1 + posY), Cells(1 + posX, 1 + posY + nroCenAGC)).Interior.Color = RGB(170, 170, 170)
    ThisWorkbook.Worksheets("Servicio AGC").Range(Cells(2 + posX, 1 + posY), Cells(2 + posX, 1 + posY + nroCenAGC)).Interior.Color = RGB(200, 200, 200)
    ThisWorkbook.Worksheets("Servicio AGC").Range(Cells(3 + posX, 1 + posY), Cells(1 + posX + 26, 1 + posY + nroCenAGC)).Interior.Color = RGB(230, 230, 230)
    ThisWorkbook.Worksheets("Servicio AGC").Range(Cells(1 + posX + 26, 1 + posY), Cells(1 + posX + 26, 1 + posY + nroCenAGC)).Interior.Color = RGB(200, 200, 200)

    For Hora = 1 To 24
        ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2 + posX, 1 + posY).Value = Hora
        ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2 + posX, 1 + posY).Interior.Color = RGB(200, 200, 200)
    Next Hora
    ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2 + posX, 1 + posY).Value = "TOTAL"
    
    For j = 1 To nroCenAGC
        ThisWorkbook.Worksheets("Servicio AGC").Cells(2 + posX, j + 1 + posY).Value = EliminarComillas(DispAGCP(j).Planta)
        TotalCentral = 0
        For Hora = 1 To 24
            ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2 + posX, j + 1 + posY).Value = DispAGCP(j).MWh(Hora)
            TotalCentral = TotalCentral + DispAGCP(j).MWh(Hora)
        Next Hora
        ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2 + posX, j + 1 + posY).Value = TotalCentral
    Next j
    
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerOfeiAgcPlanta"
    
End Sub

Sub LeerOfeiCombPlanta(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoOFEI(fecha)
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 2 Then
                If Trim(UCase(LArray(1))) = "C" Then
                    'Debug.Print textline
                    CombPlanta(i).Planta = Trim(LArray(0))
                    CombPlanta(i).comb = Trim(LArray(2))
                    i = i + 1
                End If
            End If
        Loop
    Close #1
End Sub

Sub LeerOfeiPruebas(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoOFEI(fecha)
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 25 Then
                If Trim(UCase(LArray(1))) = "PRU" Then
                    Debug.Print textline
                    PruebasPlanta(i).recurso = Trim(LArray(0))
                    For Hora = 1 To 24
                        PruebasPlanta(i).prueba(Hora) = LArray(Hora + 1)
                    Next Hora
                    i = i + 1
                End If
            End If
        Loop
    Close #1
End Sub

Sub LeerOfeiConfigPlanta(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoOFEI(fecha)
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 2 Then
                If Trim(UCase(LArray(1))) = "CONF" Then
                    'Debug.Print textline
                    ConfigPlanta(i).Planta = Trim(LArray(0))
                    ConfigPlanta(i).config = Trim(LArray(2))
                    'Debug.Print ConfigPlanta(i).planta, ConfigPlanta(i).config
                    i = i + 1
                End If
            End If
        Loop
    Close #1
End Sub
Sub LeerOfeiMinOpePlanta(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoOFEI(fecha)
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 25 Then
                If Trim(UCase(LArray(1))) = "MO" Then
                    'Debug.Print textline
                    MinOpePlanta(i).Planta = Trim(LArray(0))
                    For Hora = 1 To 24
                        MinOpePlanta(i).MWh(Hora) = LArray(Hora + 1)
                    Next Hora
                    i = i + 1
                End If
            End If
        Loop
    Close #1
End Sub

Public Sub LeerOFEI(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim fila As Integer
    Dim Hora As Integer
    Dim Total As Single
    Dim nroCampos As Integer
    Dim Valor As Single
    
    Application.Calculation = xlCalculationManual
    
    ThisWorkbook.Worksheets("OFEI").UsedRange.Delete
    
    archivo = ArchivoOFEI(fecha)
    ThisWorkbook.Worksheets("OFEI").Cells(1, 1).Value = "OFEI  " & fecha

    For Hora = 1 To 24
        ThisWorkbook.Worksheets("OFEI").Cells(1, Hora + 2).Value = "Hora " & CStr(Hora)

    Next Hora
    ThisWorkbook.Worksheets("OFEI").Cells(1, 27).Value = "Total"

    fila = 2
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            textline = Trim(textline)
            LArray = Split(textline, ",")
            nroCampos = UBound(LArray)
            If textline <> "" Then
                Total = 0
                For Hora = 0 To nroCampos
                    ThisWorkbook.Worksheets("OFEI").Cells(fila, Hora + 1).Value = LArray(Hora)
                    If nroCampos = 25 And Hora > 1 Then
                        Total = Total + LArray(Hora)
                    End If

                Next Hora
                If nroCampos = 25 Then
                    ThisWorkbook.Worksheets("OFEI").Cells(fila, 27).Value = Total
                    ThisWorkbook.Worksheets("OFEI").Cells(fila, 27).NumberFormat = "###,###,##0.00"
                End If
                fila = fila + 1
            End If
        Loop
    Close #1
    
    FormatoSimpleHoja "OFEI"

    
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerOFEI"
End Sub



Public Function ArchivoOFEI(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamOFEI, ColParamPrefijo).Value
    año = Year(fecha)
    mes = NombreMes(nmMes.Corto, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoOFEI = raiz & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamOFEI, ColParamRaiz).Value
        ArchivoOFEI = raiz & año & "\Ofe" & mes & "\" & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    End If
End Function


Public Function LeerDispTotalUnidadOFEI(nmUni As String) As Single
    Dim fila As Integer
    Dim blnHallado As Boolean
    Dim nmAux As String
    Dim strCodigo As String
    
    LeerDispTotalUnidadOFEI = 0
    nmUni = UCase(Trim(nmUni))
    
    fila = 2
    blnHallado = False
    nmAux = UCase(Trim(ThisWorkbook.Worksheets("OFEI").Cells(fila, 1).Value))
    strCodigo = UCase(Trim(ThisWorkbook.Worksheets("OFEI").Cells(fila, 2).Value))
    Do While fila < 1000 And blnHallado = False
        DoEvents
        If nmAux = nmUni And strCodigo = "D" Then
            blnHallado = True
            LeerDispTotalUnidadOFEI = ThisWorkbook.Worksheets("OFEI").Cells(fila, 27).Value
        End If
        fila = fila + 1
        nmAux = UCase(Trim(ThisWorkbook.Worksheets("OFEI").Cells(fila, 1).Value))
        strCodigo = UCase(Trim(ThisWorkbook.Worksheets("OFEI").Cells(fila, 2).Value))
     Loop
End Function