




Private Sub cmdBalances_Click()
    cmdBalances.Enabled = False
    Ejecutar_Balances
    cmdBalances.Enabled = True
End Sub

Private Sub cmdBalMañana_Click()
    ThisWorkbook.Worksheets("BalanceMañana").Activate
End Sub

Private Sub cmdCopiarEnRutaAlterna_Click()
    cmdCopiarEnRutaAlterna.Enabled = False
    CopiarEnRutaAlterna FechaReal
    cmdCopiarEnRutaAlterna.Enabled = True
End Sub

Private Sub cmdDDEC_Click()
    Dim mensaje As String
    NroErrores = 0
    Avance = 0
    MostrarAvance Avance
    LogOfertaEPM "DDEC", 3
    cmdDDEC.Enabled = False
    TiempoInicio = Time
    Application.ScreenUpdating = False
    LeerDDEC FechaReal
    Avance = Avance + 25
    MostrarAvance Avance
    LeerDDecGenProg FechaReal, 0
    Avance = Avance + 25
    MostrarAvance Avance
    LeerDDecGenProg FechaReal, 1
    Avance = Avance + 25
    MostrarAvance Avance
    LeerDDecGenProgHoraria FechaReal, 0
    Avance = Avance + 25
    MostrarAvance Avance
    ThisWorkbook.Worksheets("DDEC").Activate
    Application.ScreenUpdating = True
    TiempoFinal = Time
    cmdDDEC.Enabled = True
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "DDEC"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(FechaReal) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "DDEC"
    End If
End Sub

Private Sub cmdEmbalses_Click()
    ThisWorkbook.Worksheets("Embalses").Activate
End Sub

Private Sub cmdFecha_Click()
    usfFecha.Show 1
End Sub

Private Sub cmdGeneracion_Click()
    ThisWorkbook.Worksheets("Generacion").Activate
End Sub

Private Sub cmdGrafico_Click()
    Dim mensaje As String
    NroErrores = 0
    Avance = 0
    MostrarAvance Avance
    LogOfertaEPM "GRAFICO", 3
    cmdGrafico.Enabled = False
    ThisWorkbook.Worksheets("Grafico").Activate
    Application.ScreenUpdating = False
    TiempoInicio = Time
    DibujarEmbalses
    TiempoFinal = Time
    Application.ScreenUpdating = True
    Avance = 100
    MostrarAvance Avance
    cmdGrafico.Enabled = True
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "Grafico"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(FechaReal) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "Grafico"
    End If

End Sub

Private Sub cmdInforme_Click()
    ThisWorkbook.Worksheets("Informe").Activate
End Sub

Private Sub cmdIrADDEC_Click()
    ThisWorkbook.Worksheets("DDEC").Activate
End Sub

Private Sub cmdIrAEquivalencias_Click()
    ThisWorkbook.Worksheets("Equivalencias").Activate
End Sub

Private Sub cmdIrAFactorEmbalses_Click()
    ThisWorkbook.Worksheets("FactorEmbalses").Activate
End Sub

Private Sub cmdIrAGrafico_Click()
    ThisWorkbook.Worksheets("Grafico").Activate
End Sub

Private Sub cmdIrANEP_Click()
    ThisWorkbook.Worksheets("NEP").Activate
End Sub

Private Sub cmdIrAOFEI_Click()
    ThisWorkbook.Worksheets("OFEI").Activate
End Sub

Private Sub cmdIrAParametros_Click()
    ThisWorkbook.Worksheets("Parametros").Activate
End Sub

Private Sub cmdIrAPlantaUnidad_Click()
    ThisWorkbook.Worksheets("PlantaUnidad").Activate
End Sub

Private Sub cmdIrAPreIdeal_Click()
    ThisWorkbook.Worksheets("PreIdeal").Activate
End Sub

Private Sub cmdIrASEGDES_Click()
    ThisWorkbook.Worksheets("SEGDES").Activate
End Sub

Private Sub cmdIrASensibilidad_Click()
    ThisWorkbook.Worksheets("Sensibilidades").Activate
End Sub

Private Sub cmdIrBalance_Click()
    ThisWorkbook.Worksheets("Balances").Activate
End Sub

Private Sub cmdOFEI_Click()
    Dim mensaje As String
    NroErrores = 0
    Avance = 0
    MostrarAvance Avance
    LogOfertaEPM "OFEI", 3
    cmdOFEI.Enabled = False
    TiempoInicio = Time
    LeerOFEI FechaReal
    Avance = 100
    MostrarAvance Avance
    ThisWorkbook.Worksheets("OFEI").Activate
    TiempoFinal = Time
    cmdOFEI.Enabled = True
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "OFEI"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(FechaReal) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "OFEI"
    End If
End Sub

Private Sub cmdOfertaEPM_Click()
    Ejecutar_OfertaEPM
End Sub

Private Sub cmdOfertas_Click()
    ThisWorkbook.Worksheets("Ofertas").Activate
End Sub

Private Sub cmdPreciosGeneraciones_Click()
    ThisWorkbook.Worksheets("Precios Generaciones").Activate
End Sub

Private Sub cmdPreIdeal_Click()
    Dim mensaje As String
    NroErrores = 0
    Avance = 0
    MostrarAvance Avance
    LogOfertaEPM "PREIDEAL", 3
    cmdPreIdeal.Enabled = False
    TiempoInicio = Time
    LeerTxtPreIdeal FechaReal
    TiempoFinal = Time
    Avance = 100
    MostrarAvance Avance
    cmdPreIdeal.Enabled = True
    ThisWorkbook.Worksheets("PreIdeal").Activate
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "Preideal"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(FechaReal) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "PreIdeal"
    End If
End Sub

Private Sub cmdProgramadoReal_Click()
    ThisWorkbook.Worksheets("Programado_Real").Activate
End Sub

Private Sub cmdRevisarPreciosMarginal_Click()
    cmdRevisarPreciosMarginal.Enabled = False
    RevisarPrecioMarginal FechaReal
    cmdRevisarPreciosMarginal.Enabled = True
    Exit Sub
End Sub

Private Sub cmdRios_Click()
    ThisWorkbook.Worksheets("Rios").Activate
End Sub

Private Sub cmdSEGDES_Click()
    Dim mensaje As String
    NroErrores = 0
    Avance = 0
    MostrarAvance Avance
    LogOfertaEPM "SEGDES", 3
    cmdSEGDES.Enabled = False
    TiempoInicio = Time
    LeerSEGDES FechaReal
    ThisWorkbook.Worksheets("SEGDES").Activate
    TiempoFinal = Time
    cmdSEGDES.Enabled = True
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "SEGDES"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(FechaReal) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "SEGDES"
    End If
End Sub

Private Sub cmdSensibilidades_Click()
    Dim mensaje As String
    NroErrores = 0
    Avance = 0
    MostrarAvance Avance
    LogOfertaEPM "SENSIBILIDADES", 3
    cmdSensibilidades.Enabled = False
    TiempoInicio = Time
    Ejecutar_Sensibilidades
    TiempoFinal = Time
    Avance = 100
    MostrarAvance Avance
    cmdSensibilidades.Enabled = True
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "Sensibilidades"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(FechaReal) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "Sensibilidades"
    End If
End Sub

Private Sub cmdServicioAGC_Click()
    ThisWorkbook.Worksheets("Servicio AGC").Activate
End Sub

Private Sub cmdTableros_Click()
    cmdTableros.Enabled = False
    Ejecutar_Tableros
    cmdTableros.Enabled = True
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value = Round(DTPicker1.Value, 0)
End Sub


Private Sub DTPicker1_Change()
    ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value = Round((DTPicker1.Value), 0)
    txtFecha.Text = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value
End Sub






Private Sub optSI_RAlt_Click()
    If optSI_RAlt.Value = True Then
        optNO_RAlt.Value = False
        blnCopiarEnRutaAlterna = True
        ThisWorkbook.Worksheets("Parametros").Cells(FilaParamCopiarEnRutaAlterna, ColParamUsarRutaAlterna).Value = True
    End If
End Sub
Private Sub optNO_RAlt_Click()
    If optNO_RAlt.Value = True Then
        optSI_RAlt.Value = False
        blnCopiarEnRutaAlterna = False
        ThisWorkbook.Worksheets("Parametros").Cells(FilaParamCopiarEnRutaAlterna, ColParamUsarRutaAlterna).Value = False
    End If
End Sub

Private Sub optSi_Click()
    If optSi.Value = True Then
        optNo.Value = False
        blnUsarRutaAlterna = True
        ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamUsarRutaAlterna).Value = True
        cmdCopiarEnRutaAlterna.Visible = False
        frmCopiaEnRutaAlterna.Visible = False
        optNO_RAlt.Value = True
    End If
End Sub
Private Sub optNo_Click()
    If optNo.Value = True Then
        optSi.Value = False
        blnUsarRutaAlterna = False
        ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamUsarRutaAlterna).Value = False
        cmdCopiarEnRutaAlterna.Visible = True
        frmCopiaEnRutaAlterna.Visible = True
    End If
End Sub


Private Sub txtCasoOfertado_Change()
    If IsNumeric(txtCasoOfertado.Text) Then
        ThisWorkbook.Worksheets("Sensibilidades").Cells(FilaSenCasoOfertado, ColSenCasoOfertado).Value = txtCasoOfertado.Text
        casoOfertado = ThisWorkbook.Worksheets("Sensibilidades").Cells(FilaSenCasoOfertado, ColSenCasoOfertado).Value
    Else
        Beep
        txtCasoOfertado.SetFocus
    End If
End Sub


Private Sub txtFecha_Change()
    Dim sTexto As String
    On Error GoTo Salida
    sTexto = Trim(txtFecha.Text)
    If Len(sTexto) = 10 Then
        sTexto = Format(sTexto, FormatoFechaWindows)
        If IsDate(sTexto) Then
           ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value = CDate(sTexto)
           FechaReal = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value
        Else
            Beep
            txtFecha.SetFocus
        End If
    End If
    Exit Sub
Salida:
    Beep
    On Error GoTo 0
    Exit Sub
End Sub


Private Sub UserForm_Activate()
    On Error Resume Next
    txtFecha.Text = CDate(ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value)
    txtFecha.ControlTipText = FormatoFechaWindows
    lblFecha.ControlTipText = FormatoFechaWindows
    txtCasoOfertado.Text = ThisWorkbook.Worksheets("Sensibilidades").Cells(FilaSenCasoOfertado, ColSenCasoOfertado).Value
    txtCaso.Text = UltimoCaso
    If blnUsarRutaAlterna Then optSi.Value = True Else optNo.Value = True
    If blnCopiarEnRutaAlterna Then optSI_RAlt.Value = True Else optNO_RAlt.Value = True
    Caption = "Oferta EPM " & VersionOfertaEPM
End Sub


