Option Explicit
Public Const ClaveVBA As String = "ofertaepmv100epm"
Public strLogin As String
Public OfertaEPMRibbon As IRibbonUI
Public strCaptionOriginal As String

Public TiempoInicio As Double
Public TiempoFinal As Double
Public VersionModelo As String
Public VersionModeloAnterior As String
Public VersionParaBorrar As String
Public Const VersionOfertaEPM As String = "V 1 1 0"
Public Const Empresa As String = "Empresas Públicas de Medellín E.S.P."

Public Const Soporte As String = "Unidad Soluciones TI Comerciales - José Fernando Bosch Moreno"

'   "EPM"
Public Const Esquema As String = "EPM"
Public Const Licenciado As String = "Gerencia Mercado Energia Mayorista"
Public Const Sigla As String = "Unidad Gestión Bolsa Energía"

#If VBA7 Then
    #If Win64 Then
        Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    #Else
        Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    #End If
#Else
    Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

Public Sub CerrarPresentacion()
    Unload usfPresentacion
End Sub

Function Usuario() As String
    Dim lSize As Long
    Dim sBuffer As String
    Dim gLoginUsuario As String
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
       gLoginUsuario = UCase(Left$(sBuffer, lSize - 1))
    Else
       gLoginUsuario = vbNullString
    End If
    strLogin = gLoginUsuario
    Usuario = strLogin
End Function

Public Function UsuarioValido(nmUsuario As String) As Boolean
    UsuarioValido = False
    nmUsuario = UCase(Trim(nmUsuario))
    If nmUsuario = UCase("aaristiz") Then UsuarioValido = True
    If nmUsuario = UCase("aospinab") Then UsuarioValido = True
    If nmUsuario = UCase("cgaviria") Then UsuarioValido = True
    If nmUsuario = UCase("czuluaga") Then UsuarioValido = True
    If nmUsuario = UCase("dgomezv") Then UsuarioValido = True
    If nmUsuario = UCase("drodrigm") Then UsuarioValido = True
    If nmUsuario = UCase("emayas") Then UsuarioValido = True
    If nmUsuario = UCase("falvarea") Then UsuarioValido = True
    If nmUsuario = UCase("fbedoyav") Then UsuarioValido = True
    If nmUsuario = UCase("groldana") Then UsuarioValido = True
    If nmUsuario = UCase("jbosch") Then UsuarioValido = True
    If nmUsuario = UCase("jhernm") Then UsuarioValido = True
    If nmUsuario = UCase("jortizr") Then UsuarioValido = True
    If nmUsuario = UCase("lpuerta") Then UsuarioValido = True
    If nmUsuario = UCase("lrendonp") Then UsuarioValido = True
    If nmUsuario = UCase("nbustama") Then UsuarioValido = True
    If nmUsuario = UCase("pchinchi") Then UsuarioValido = True
    If nmUsuario = UCase("rcalled") Then UsuarioValido = True
    If nmUsuario = UCase("rlondonr") Then UsuarioValido = True
End Function


Sub Clave()
    MsgBox ClaveVBA
End Sub


Public Function esUsuarioValido(pUsuario As String) As Boolean
    esUsuarioValido = False
    If Trim(UCase(pUsuario)) = "JBOSCH" Then esUsuarioValido = True
    If Trim(UCase(pUsuario)) = "LPUERTA" Then esUsuarioValido = True
    If Trim(UCase(pUsuario)) = "MRODRIGE" Then esUsuarioValido = True
    If Trim(UCase(pUsuario)) = "ECARVAJA" Then esUsuarioValido = True
    If Trim(UCase(pUsuario)) = "SVELEZGO" Then esUsuarioValido = True
End Function
Sub rbOnLoad(ribbon As IRibbonUI)
    Set OfertaEPMRibbon = ribbon
    If CInt(Application.Version) > 12 Then
        OfertaEPMRibbon.ActivateTab ("OfertaEPMTab")
    End If
End Sub


Sub Ejecutar_MenuPrincipal()
    MenuPrincipal
End Sub

Sub MenuPrincipal()
    usfPrincipal.Show 0
End Sub

Sub rbMenuPrincipal(ByVal control As IRibbonControl)
    MenuPrincipal
End Sub

Sub rbEjecutarTableros(ByVal control As IRibbonControl)
    Ejecutar_Tableros
End Sub

Sub rbIrOfertas(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Ofertas").Activate
End Sub

Sub rbIrProgramadoReal(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Programado_Real").Activate
End Sub

Sub rbIrPreciosGeneraciones(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Precios Generaciones").Activate
End Sub

Sub rbIrServicioAGC(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Servicio AGC").Activate
End Sub

Sub rbEjecutarSensibilidades(ByVal control As IRibbonControl)
    Ejecutar_Sensibilidades
End Sub

Sub rbIrSensibilidades(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Sensibilidades").Activate
End Sub

Sub rbEjecutarBalances(ByVal control As IRibbonControl)
    Ejecutar_Balances
End Sub

Sub rbEjecutarGrafico(ByVal control As IRibbonControl)
    DibujarEmbalses
End Sub

Sub rbIrEmbalses(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Embalses").Activate
End Sub

Sub rbIrNEP(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("NEP").Activate
End Sub


Sub rbIrGrafico(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Grafico").Activate
End Sub

Sub rbIrGeneracion(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Generacion").Activate
End Sub

Sub rbIrRios(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Rios").Activate
End Sub

Sub rbIrFactores(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("FactorEmbalses").Activate
End Sub
Sub rbIrBalances(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Balances").Activate
End Sub
Sub rbIrBalMañana(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("BalanceMañana").Activate
End Sub
Sub rbIrInforme(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("Informe").Activate
End Sub

Sub rbIrPreIdeal(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("PreIdeal").Activate
End Sub

Sub rbIrDDEC(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("DDEC").Activate
End Sub

Sub rbIrOFEI(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("OFEI").Activate
End Sub

Sub rbIrSEGDES(ByVal control As IRibbonControl)
    ThisWorkbook.Worksheets("SEGDES").Activate
End Sub

Sub rbAcercaDe(ByVal control As IRibbonControl)
    usfPresentacion.Show 0
End Sub


Sub rbCalcularPresiones(ByVal control As IRibbonControl)
    MsgBox "Opcion no implementada"
    'EjecutarCalcularPresiones
End Sub

Sub rbCalcularRestricciones(ByVal control As IRibbonControl)
    MsgBox "Opcion no implementada"
    'EjecutarCalcularRestricciones
End Sub

Sub rbVerPresiones(ByVal control As IRibbonControl)
    MsgBox "Opcion no implementada"
    'EjecutarVerPresiones
End Sub

Sub rbVerConsumos(ByVal control As IRibbonControl)
    MsgBox "Opcion no implementada"
    'EjecutarVerConsumos
End Sub

Sub Ejecutar_OfertaEPM()
    FechaReal = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value
    InformeOfertaEPM FechaReal
End Sub

Sub Ejecutar_Tableros()
    FechaReal = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value
    InformeTableros FechaReal
End Sub

Sub Ejecutar_Sensibilidades()
    FechaReal = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value
    LeerArchivoSensibilidadesMobe FechaReal + 1
End Sub

Sub Ejecutar_Balances()
    FechaReal = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamFechaReal, ColParamFechaReal).Value
    InformeBalance FechaReal
End Sub

Sub MostrarAvance(porc As Integer)
    'usfPrincipal.lblPorcAvance.Width = usfPrincipal.lblPorcAvance.Width + porc * usfPrincipal.frProgressBar.Width / 100
    usfPrincipal.lblPorcAvance.Width = porc * (usfPrincipal.frProgressBar.Width - 14) / 100
    usfPrincipal.lblPorcAvance.Caption = porc & " %"
    
End Sub


Function FechaCompilacion() As Date
    FechaCompilacion = "2019/02/06"
End Function


Sub FormatoBalance()
    Dim i As Integer
    For i = 3 To 358
        If Application.WorksheetFunction.IsFormula(ThisWorkbook.Worksheets("Balances").Cells(i, 6)) _
           Or IsEmpty(ThisWorkbook.Worksheets("Balances").Cells(i, 6)) Then
            ThisWorkbook.Worksheets("Balances").Cells(i, 6).Interior.ThemeColor = xlThemeColorDark1
            ThisWorkbook.Worksheets("Balances").Cells(i, 6).Interior.TintAndShade = -4.99893185216834E-02
        Else
            ThisWorkbook.Worksheets("Balances").Cells(i, 6).Interior.Color = 65535
        End If
        
        If Application.WorksheetFunction.IsFormula(ThisWorkbook.Worksheets("Balances").Cells(i, 7)) _
           Or IsEmpty(ThisWorkbook.Worksheets("Balances").Cells(i, 7)) Then
            ThisWorkbook.Worksheets("Balances").Cells(i, 7).Interior.ThemeColor = xlThemeColorDark1
            ThisWorkbook.Worksheets("Balances").Cells(i, 7).Interior.TintAndShade = -4.99893185216834E-02
        Else
            ThisWorkbook.Worksheets("Balances").Cells(i, 7).Interior.Color = 65535
        End If
        
        If (ThisWorkbook.Worksheets("Balances").Cells(i, 1).Value = "Volumen inicial" Or _
            ThisWorkbook.Worksheets("Balances").Cells(i, 1).Value = "Volumen Final") And _
            ThisWorkbook.Worksheets("Balances").Cells(i, 4).Value = "%" Then
            ThisWorkbook.Worksheets("Balances").Cells(i, 6).Interior.ThemeColor = xlThemeColorAccent1
            ThisWorkbook.Worksheets("Balances").Cells(i, 6).Interior.TintAndShade = 0.599993896298105
            ThisWorkbook.Worksheets("Balances").Cells(i, 7).Interior.ThemeColor = xlThemeColorAccent1
            ThisWorkbook.Worksheets("Balances").Cells(i, 7).Interior.TintAndShade = 0.599993896298105
        End If
    
    Next i
End Sub


Sub ProtegerFormulas()
    Dim i As Integer
    For i = 3 To 358
        If Application.WorksheetFunction.IsFormula(ThisWorkbook.Worksheets("Balances").Cells(i, 6)) _
           Or IsEmpty(ThisWorkbook.Worksheets("Balances").Cells(i, 6)) Then
            'ThisWorkbook.Worksheets("Balances").Cells(i, 6).Interior.ThemeColor = xlThemeColorDark1
            'ThisWorkbook.Worksheets("Balances").Cells(i, 6).Interior.TintAndShade = -4.99893185216834E-02
        Else
            ThisWorkbook.Worksheets("Balances").Cells(i, 6).Locked = False
        End If
    
    Next i
End Sub


Sub protegerHoja(nmHoja As String, Optional protegerHoja As Boolean)
    Dim i As Integer
    Dim c As Range
    Application.ScreenUpdating = False
    TiempoInicio = Time
    For Each c In ThisWorkbook.Worksheets(nmHoja).UsedRange.Cells
        If Application.WorksheetFunction.IsFormula(c) _
           Or IsEmpty(c) Then
            c.Locked = True
        Else
            c.Locked = False
        End If
    Next
    ThisWorkbook.Worksheets(nmHoja).Protect Userinterfaceonly:=True
    Application.ScreenUpdating = True
    TiempoFinal = Time
    MsgBox "Fin " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 2))
End Sub


Sub ProtegerHojas()
    protegerHoja "Parametros"
    protegerHoja "PlantaUnidad"
    protegerHoja "Equivalencias"
    protegerHoja "Ofertas"
    protegerHoja "Precios Generaciones"
    protegerHoja "Programado_Real"
    protegerHoja "Servicio AGC"
    protegerHoja "Sensibilidades"
    protegerHoja "Preideal"
    protegerHoja "DDEC"
    protegerHoja "OFEI"
    protegerHoja "SEGDES"
    protegerHoja "nep"
    protegerHoja "Rios"
    protegerHoja "Embalses"
    protegerHoja "Generacion"
    protegerHoja "FactorEmbalses"
    protegerHoja "Balances"
    protegerHoja "BalanceMañana"
    protegerHoja "Informe"
End Sub
