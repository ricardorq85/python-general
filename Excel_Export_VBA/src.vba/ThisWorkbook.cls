
Private Sub Workbook_Open()
    If UsuarioValido(Usuario()) Then
        Application.OnKey "^%{m}", "Ejecutar_MenuPrincipal"
        FechaReal = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value
        Application.OnKey "+^%&", "Clave"
        Application.OnTime Now + TimeValue("0:00:04"), "CerrarPresentacion", , True
        usfPresentacion.Show False
    Else
        Beep
        MsgBox "Usuario no valido", vbOKOnly, "OfertaEPM"
        'ThisWorkbook.Close False
    End If
End Sub

