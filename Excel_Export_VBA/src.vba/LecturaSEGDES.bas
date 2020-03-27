Public Sub LeerSEGDES(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim fila As Integer
    Dim Hora As Integer
    Dim Total As Single
    Dim nroCampos As Integer
    Dim Valor As Single
    
    Application.Calculation = xlCalculationManual
    
    ThisWorkbook.Worksheets("SEGDES").UsedRange.Delete
    
    archivo = ArchivoDSEGDES(fecha)
    ThisWorkbook.Worksheets("SEGDES").Cells(1, 1).Value = "SEGDES  " & fecha

    For Hora = 1 To 24
        ThisWorkbook.Worksheets("SEGDES").Cells(1, Hora + 2).Value = "Hora " & CStr(Hora)

    Next Hora
    ThisWorkbook.Worksheets("SEGDES").Cells(1, 27).Value = "Total"

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
                    ThisWorkbook.Worksheets("SEGDES").Cells(fila, Hora + 1).Value = LArray(Hora)
                    If nroCampos = 25 And Hora > 1 Then
                        Total = Total + LArray(Hora)
                    End If

                Next Hora
                If nroCampos = 25 Then
                    ThisWorkbook.Worksheets("SEGDES").Cells(fila, 27).Value = Total
                    ThisWorkbook.Worksheets("SEGDES").Cells(fila, 27).NumberFormat = "###,###,##0.00"
                End If
                fila = fila + 1
            End If
        Loop
    Close #1
    
    FormatoSimpleHoja "SEGDES"

    
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerSEGDES"
End Sub

Sub CargarSEGDES_DispCen(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoDSEGDES(fecha)
    Dim Hora As Integer
    Dim nmCen As String
    Dim nmUni As String
    Dim codHashCen As Integer
    Dim Valor As Single
    'nroCenAGC = 0
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 25 Then
                If EliminarComillas(UCase(Trim(LArray(1)))) = "MX" Then
                    nmUni = EliminarComillas(UCase(Trim(LArray(0))))
                    nmCen = UCase(Trim(nmCentralUnidad(nmUni, CentralDeUnidad, nroUnidades)))
                    codHashCen = CodigoHash(nmCen)
                    For Hora = 1 To 24
                        Valor = LArray(Hora + 1)
                        DispCen(codHashCen).MX(Hora) = DispCen(codHashCen).MX(Hora) + Valor
                    Next Hora
                    'Debug.Print codHashCen; "  "; textline
                End If
                If EliminarComillas(UCase(Trim(LArray(1)))) = "MW" Then
                    nmUni = EliminarComillas(UCase(Trim(LArray(0))))
                    nmCen = UCase(Trim(nmCentralUnidad(nmUni, CentralDeUnidad, nroUnidades)))
                    codHashCen = CodigoHash(nmCen)
                    For Hora = 1 To 24
                        Valor = LArray(Hora + 1)
                        DispCen(codHashCen).MW(Hora) = DispCen(codHashCen).MW(Hora) + Valor
                    Next Hora
                    'Debug.Print codHashCen; "  "; textline
                End If
            End If
        Loop
    Close #1
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " CargarSEGDES_DispCen"
End Sub

Public Function ArchivoDSEGDES(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamSEGDES, ColParamPrefijo).Value
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoDSEGDES = raiz & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamSEGDES, ColParamRaiz).Value
        ArchivoDSEGDES = raiz & año & "\" & mes & "\" & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    End If
End Function