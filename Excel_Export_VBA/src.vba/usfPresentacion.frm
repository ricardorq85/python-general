






Private Sub UserForm_Initialize()
    lblEmpresa.Caption = Empresa
    lblVersion.Caption = VersionOfertaEPM & " " & CStr(FechaCompilacion)
    lblLicenciado.Caption = Licenciado
    lblSigla.Caption = Sigla
    lblNombreBD = BaseDatos
    lblNombreEsquema = Esquema
    lblNombreDSN = DSN
    lblNombreSoporte = Soporte
    lblNombreServidor = NombreServidor
    lblUsuario.Caption = "Usuario: " & Usuario()
End Sub