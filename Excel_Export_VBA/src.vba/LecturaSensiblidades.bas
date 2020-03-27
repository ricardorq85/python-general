Option Explicit

Public vLongFileSensibilidad As Variant
Public vTimeFileSensibilidad As Variant
Public bSeEstaEjecutandoProceso As Boolean

'Sub ProcGenerarSensibilidadesMobe()
'    If bSeEstaEjecutandoProceso = False Then
'        bSeEstaEjecutandoProceso = True
'        LeerArchivoSensibilidadesMobe FechaREal
'        While ActiveSheet.Name = "Sensibilidades"
'            DoEvents
'            If Int(Timer) = Int(Timer / 5) * 5 Then
'                LeerArchivoSensibilidadesMobe FechaREal
'            End If
'        Wend
'        bSeEstaEjecutandoProceso = False
'    End If
'End Sub


Sub LeerArchivoSensibilidadesMobe(fecha As Date)
    Dim ArchivoSensibilidad As String
    Dim sLineaTexto As String
    Dim sLineaAux As String
    Dim iNroColum As Integer
    Dim iNroValFila As Integer
    Dim iNroValFilaAnt As Integer
    Dim iPosFilIni As Integer
    Dim iPosFilFin As Integer
    Dim vCeldaActiva As Variant
    Dim iNroAuxiliar As Integer
    Dim fhReporteSensible As Date

    Dim iValor As Integer
    
    fhReporteSensible = Date + 1  'Nota se agrega para revisar
    
    On Error GoTo ManejadorError
    
    ArchivoSensibilidad = ArchivoInfSen(fecha)
    
    If Len(Dir(ArchivoSensibilidad, vbArchive)) > 0 Then
        'If (vLongFileSensibilidad <> FileLen(ArchivoSensibilidad)) Or _
        '        (vTimeFileSensibilidad <> FileDateTime(ArchivoSensibilidad)) Then

            ThisWorkbook.Worksheets("Sensibilidades").Range("A2:AA200").Delete
            With Selection
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            iNroValFilaAnt = 1
            iPosFilIni = 2
            ThisWorkbook.Worksheets("Sensibilidades").Activate
            ThisWorkbook.Worksheets("Sensibilidades").Range("$B$2").Activate
            
            Open ArchivoSensibilidad For Input As #1
                ThisWorkbook.Worksheets("Sensibilidades").Range("$B$2").Activate
            
            Do While Not EOF(1)
                Line Input #1, sLineaTexto
                If (InStr(1, sLineaTexto, "MANTIOQ1") > 0) Or _
                        (InStr(1, sLineaTexto, "MJEPIRACHI") > 0) Then
                    sLineaTexto = " "
                Else
                    If (InStr(1, sLineaTexto, "PRECIOS OFERTAS") > 0) Then
                        ActiveCell.Offset(0, -1).Select
                        ActiveCell.Value = "PRECIOS OFERTAS"
                        ActiveCell.Offset(0, 1).Select
                        sLineaTexto = " "
                    ElseIf (InStr(1, sLineaTexto, "RESULTADOS EPM") > 0) Then
                        ActiveCell.Offset(0, -1).Select
                        ActiveCell.Value = "EPM"
                        ActiveCell.Offset(0, 1).Select
                        sLineaTexto = " "
                    ElseIf (InStr(1, sLineaTexto, "GENERACION REAL PLANTAS") > 0) Then
                        ActiveCell.Offset(0, -1).Select
                        ActiveCell.Value = "GENERACION REAL PLANTAS"
                        ActiveCell.Offset(0, 1).Select
                        sLineaTexto = " "
                    ElseIf (InStr(1, sLineaTexto, "SERVICIO AGC") > 0) Then
                        ActiveCell.Offset(0, -1).Select
                        ActiveCell.Value = "SERVICIO AGC"
                        ActiveCell.Offset(0, 1).Select
                        sLineaTexto = " "
                    ElseIf (InStr(1, sLineaTexto, "GENERACION REAL POR EMPRESA") > 0) Then
                        ActiveCell.Offset(0, -1).Select
                        ActiveCell.Value = "GENERACION REAL POR EMPRESA"
                        ActiveCell.Offset(0, 1).Select
                        sLineaTexto = " "
                    ElseIf (InStr(1, sLineaTexto, "GENERACION IDEAL POR EMPRESA") > 0) Then
                        ActiveCell.Offset(0, -1).Select
                        ActiveCell.Value = "GENERACION IDEAL POR EMPRESA"
                        ActiveCell.Offset(0, 1).Select
                        sLineaTexto = " "
                    ElseIf (InStr(1, sLineaTexto, "PRECIO DE BOLSA HORARIO EMPRESA") > 0) Then
                        ActiveCell.Offset(0, -1).Select
                        ActiveCell.Value = "PRECIO DE BOLSA HORARIO EMPRESA"
                        ActiveCell.Offset(0, 1).Select
                        sLineaTexto = " "
                    ElseIf (InStr(1, sLineaTexto, "DIFERENCIA HORARIA DESPACHO REAL CONTRATOS EPM (MWh)") > 0) Then
                        ActiveCell.Offset(0, -1).Select
                        ActiveCell.Value = "DIFERENCIA HORARIA DESPACHO REAL CONTRATOS EPM (MWh)"
                        ActiveCell.Offset(0, 1).Select
                        sLineaTexto = " "
                    End If
                End If
                If Len(Trim(sLineaTexto)) > 0 Then
                    iNroValFila = Mid(sLineaTexto, 1, 1)
                    If iNroValFilaAnt <> iNroValFila Then
                        ' linea en blanco
                        iPosFilFin = ActiveCell.Cells.Row - 1
                        vCeldaActiva = ActiveCell.Address
                        If iPosFilFin > 0 Then
                            iNroAuxiliar = funcOrganizarPresentacionCeldasSensibilidades(iPosFilIni, iPosFilFin, iNroColum, iNroValFilaAnt)
                        Else
                            iNroAuxiliar = 0
                        End If
                        iNroValFilaAnt = iNroValFila
                        iPosFilIni = iPosFilFin + 1
                        ThisWorkbook.Worksheets("Sensibilidades").Range(vCeldaActiva).Activate
                    End If
                    sLineaAux = Trim(Mid(sLineaTexto, InStr(1, sLineaTexto, "|") + 1))
                    iNroColum = 0
                    While InStr(1, sLineaAux, "|") > 0
                        iNroColum = iNroColum + 1
                        ActiveCell.Value = Trim(UCase(Mid(sLineaAux, 1, InStr(1, sLineaAux, "|") - 1)))
                        ActiveCell.Offset(0, 1).Select
                        sLineaAux = Mid(sLineaAux, InStr(1, sLineaAux, "|") + 1)
                    Wend
                    ActiveCell.Value = Trim(UCase(sLineaAux))
                    ActiveCell.Offset(1, -iNroColum).Select
                End If
            Loop
            Close #1
            iPosFilFin = ActiveCell.Cells.Row - 1
            ActiveCell.Offset(1, 0).Select
            vCeldaActiva = ActiveCell.Address
            If iPosFilFin > 0 Then
                iNroAuxiliar = funcOrganizarPresentacionCeldasSensibilidades(iPosFilIni, iPosFilFin, iNroColum, iNroValFilaAnt)
            Else
                iNroAuxiliar = 0
            End If
            
            
            ThisWorkbook.Worksheets("Sensibilidades").Range("A2:A100").Select
            With Selection.Font
                  .Name = "Arial"
                  .Size = 8
                  .Strikethrough = False
                  .Superscript = False
                  .Subscript = False
                  .OutlineFont = False
                  .Shadow = False
                  .Underline = xlUnderlineStyleNone
                  .ColorIndex = xlAutomatic
                  '.TintAndShade = 0
                  '.ThemeFont = xlThemeFontNone
            End With

            
            ThisWorkbook.Worksheets("Sensibilidades").Range("B1").Activate
            vLongFileSensibilidad = FileLen(ArchivoSensibilidad)
            vTimeFileSensibilidad = FileDateTime(ArchivoSensibilidad)
            ThisWorkbook.Worksheets("Sensibilidades").Cells(2, 2).Value = fecha
            ThisWorkbook.Worksheets("Sensibilidades").Cells(FilaSenCasoOfertado, ColSenCasoOfertado).Value = UltimoCaso

        'End If
    End If
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & ArchivoSensibilidad & " LeerArchivoSensibilidadesMobe"
End Sub

Public Function ArchivoInfSen(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamInfSen, ColParamPrefijo).Value
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoInfSen = raiz & prefijo & NombreDia(nmDia.Corto, fecha) & NombreMes(nmMes.Corto, fecha) & nroDia(ConCero, fecha) & ".txt"
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamInfSen, ColParamRaiz).Value
        ArchivoInfSen = raiz & año & "\" & mes & "\Oferta\" & prefijo & NombreDia(nmDia.Corto, fecha) & NombreMes(nmMes.Corto, fecha) & nroDia(ConCero, fecha) & ".txt"
    End If
End Function


Function funcOrganizarPresentacionCeldasSensibilidades(iFilIni As Integer, _
                iFilFin As Integer, iColFin As Integer, iTipoDato As Integer)
    Dim sNmColumnFinal As String
    
    ActiveCell.Offset(-1, 0).Select
    ActiveCell.Offset(, iColFin).Select
    sNmColumnFinal = Mid(ActiveCell.Address, 2, 1)
    Rows("1:200").Select
    Selection.RowHeight = 12.75
    Columns("C:" & sNmColumnFinal).Select
    Selection.ColumnWidth = 16.5
    ThisWorkbook.Worksheets("Sensibilidades").Range("A" & iFilIni & ":" & sNmColumnFinal & iFilFin).Select
    ProcPintarCuadrillasDatos
    ThisWorkbook.Worksheets("Sensibilidades").Range("A" & iFilIni & ":" & sNmColumnFinal & iFilFin).Select
            With Selection.Font
                .Name = "Arial"
                .Size = 9
                .Bold = True
            End With
    ThisWorkbook.Worksheets("Sensibilidades").Range("A" & iFilIni & ":" & sNmColumnFinal & iFilFin).Select
    If (iTipoDato = 1) Then
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Else
        With Selection
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
        End With
    End If
    If (iTipoDato = 5) Then
        ThisWorkbook.Worksheets("Sensibilidades").Range("C" & iFilIni & ":" & sNmColumnFinal & iFilFin).Select
        With Selection
            .HorizontalAlignment = xlRight
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
    End If
    ThisWorkbook.Worksheets("Sensibilidades").Range("A" & iFilIni & ":" & sNmColumnFinal & iFilFin).Select
    If (iTipoDato = 1) Then
        With Selection.Interior
            .ColorIndex = 36
            .Pattern = xlSolid
        End With
    ElseIf (iTipoDato = 2) Then
        With Selection.Interior
            .ColorIndex = 37
            .Pattern = xlSolid
        End With
    ElseIf (iTipoDato = 3) Then
        With Selection.Interior
            .ColorIndex = 35
            .Pattern = xlSolid
        End With
    ElseIf (iTipoDato = 4) Then
        With Selection.Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With
    ElseIf (iTipoDato = 5) Then
        With Selection.Interior
            .ColorIndex = 34
            .Pattern = xlSolid
        End With
    ElseIf (iTipoDato = 6) Then
        With Selection.Interior
            .ColorIndex = 40
            .Pattern = xlSolid
        End With
    ElseIf (iTipoDato = 7) Then
        With Selection.Interior
            .ColorIndex = 38 ' 39
            .Pattern = xlSolid
        End With
    ElseIf (iTipoDato = 8) Then
        With Selection.Interior
            .ColorIndex = 42
            .Pattern = xlSolid
        End With
    ElseIf (iTipoDato = 9) Then
        With Selection.Interior
            .ColorIndex = 44
            .Pattern = xlSolid
        End With
    End If
    ThisWorkbook.Worksheets("Sensibilidades").Columns("A").ColumnWidth = 11
    
    ThisWorkbook.Worksheets("Sensibilidades").Columns("B").ColumnWidth = 23
    
    ThisWorkbook.Worksheets("Sensibilidades").Range("A" & iFilIni & ":A" & iFilFin).Select
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    If (iTipoDato = 3) Then
        ThisWorkbook.Worksheets("Sensibilidades").Range("C" & iFilIni & ":" & sNmColumnFinal & iFilIni + 1).NumberFormat = "#,##0.0; -#,##0.0"
    ElseIf (iTipoDato = 4) Then
        ThisWorkbook.Worksheets("Sensibilidades").Range("C" & iFilIni & ":" & sNmColumnFinal & iFilFin).NumberFormat = "#,##0.0; -#,##0.0"
    ElseIf (iTipoDato > 5) Then
        ThisWorkbook.Worksheets("Sensibilidades").Range("C" & iFilIni & ":" & sNmColumnFinal & iFilFin).NumberFormat = "#,##0.0; -#,##0.0"
    End If
    funcOrganizarPresentacionCeldasSensibilidades = 0
End Function

Sub ProcPintarCuadrillasDatos()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    If Selection.Columns.Count > 1 Then
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    If Selection.Rows.Count > 1 Then
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
End Sub

Public Function GenRealSen(nmCen As String, caso As Integer) As Single
    Dim fila As Integer
    Dim filaCaso As Integer
    Dim colCasoCero As Integer
    Dim nroCasos As Integer
    Dim colCaso As Integer
    Dim FilaInicioGenReal As Integer
    Dim Seccion As String
    Dim nmSeccion As String
    Dim blnHallado As Boolean
    Dim central As String
    Dim nmCentral As String
    
    
    GenRealSen = -1
    fila = 1
    filaCaso = 1
    colCasoCero = 3
    blnHallado = False
    nmSeccion = "GENERACION REAL PLANTAS"
    Seccion = UCase(Trim(ThisWorkbook.Worksheets("Sensibilidades").Cells(fila, 1).Value))
    Do While blnHallado = False And fila < 200
        DoEvents
        If Seccion = nmSeccion Then
            blnHallado = True
            FilaInicioGenReal = fila
        End If
    
        fila = fila + 1
        Seccion = UCase(Trim(ThisWorkbook.Worksheets("Sensibilidades").Cells(fila, 1).Value))
    Loop
    
    If blnHallado Then
        blnHallado = False
        nmCentral = UCase(Trim(nmCen))
        fila = FilaInicioGenReal
        central = UCase(Trim(ThisWorkbook.Worksheets("Sensibilidades").Cells(fila, 2).Value))
        
        'hallar fila central en sensibilidades
        Do While blnHallado = False And fila < 200
            DoEvents
            If central = nmCentral Then
                blnHallado = True
                GenRealSen = ThisWorkbook.Worksheets("Sensibilidades").Cells(fila, colCasoCero + caso).Value
            End If
            fila = fila + 1
            central = UCase(Trim(ThisWorkbook.Worksheets("Sensibilidades").Cells(fila, 2).Value))
            
        Loop
    Else
        GenRealSen = -1
    End If

End Function