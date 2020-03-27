Option Explicit
Public Avance As Integer
Public NroErrores As Integer
Public strError As String
Public strArchivoLog As String

Sub InformeOfertaEPM(fecha As Date)
    Dim strArchivo As String
    Dim mensaje As String
    
    NroErrores = 0
 
    On Error GoTo ManejadorError
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    ThisWorkbook.Worksheets("Parametros").Cells(1, 4).Value = Usuario()
    ThisWorkbook.Worksheets("Parametros").Cells(1, 5).Value = "Fecha: " & CStr(Date) & " " & CStr(Time)
    TiempoInicio = Time
    
    Avance = 0
    NroErrores = 0
    blnUsarRutaAlterna = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamUsarRutaAlterna).Value
    blnCopiarEnRutaAlterna = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamCopiarEnRutaAlterna, ColParamUsarRutaAlterna).Value
    If blnCopiarEnRutaAlterna Then CopiarEnRutaAlterna fecha
    
    LogOfertaEPM "Inicio OfertaEPM Versión " & VersionOfertaEPM & " para la fecha " & CStr(fecha) & ", Ejecutado el dia: " & CStr(Date) & "  " & CStr(Time()) & "  " & Usuario(), 3
    LogOfertaEPM "------OFERTAEPM--------", 3
    LogOfertaEPM "------TABLEROS---------", 3
    If Not (blnUsarRutaAlterna) Then CopiarArchivoIDOdesdeDiario fecha - 1
    
    If ExisteArchivoFuente("PRID", fecha) Then
        LeerTxtPreIdeal fecha
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM "  1 - No se encontro " & ArchivoPreideal(fecha)
    End If
    '3
    
    If ExisteArchivoFuente("DDEC", fecha) Then
        LeerDDEC fecha
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM "  2 - No se encontro " & ArchivoDDEC(fecha)
    End If
    '6
    
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabElePreciosOferta fecha, Prioridades, 0
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM "  3 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '8
        
    If ExisteArchivoFuente("INFTABELE", fecha - 1) Then
        LeerInfTabElePreciosOferta fecha, Prioridades, 1
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM "  4 - No se encontro " & ArchivoInfTabEle(fecha - 1)
    End If
    '10
        
    If ExisteArchivoFuente("INFTABELE", fecha - 2) Then
        LeerInfTabElePreciosOferta fecha, Prioridades, 2
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM "  5 - No se encontro " & ArchivoInfTabEle(fecha - 2)
    End If
    '12
        
    If ExisteArchivoFuente("INFTABELE", fecha - 3) Then
        LeerInfTabElePreciosOferta fecha, Prioridades, 3
        Avance = Avance + 4
        MostrarAvance Avance
    Else
        LogOfertaEPM "  6 - No se encontro " & ArchivoInfTabEle(fecha - 3)
    End If
    '16
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabElePrioridades fecha
        Avance = Avance + 4
        MostrarAvance Avance
    Else
        LogOfertaEPM "  7 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '20
        
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabElePrecioBolsa fecha
        Avance = Avance + 4
        MostrarAvance Avance
    Else
        LogOfertaEPM "  8 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '24
      
    If ExisteArchivoFuente("DMAR", fecha) Then
        LeerDMAR fecha
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM "  9 - No se encontro " & ArchivoDMAR(fecha)
    End If
    '27
    
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabEleResultados fecha, 0
        Avance = Avance + 4
        MostrarAvance Avance
    Else
        LogOfertaEPM " 10 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '31

    If ExisteArchivoFuente("INFTABELE", fecha - 1) Then
        LeerInfTabEleResultados fecha, 1
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 11 - No se encontro " & ArchivoInfTabEle(fecha - 1)
    End If
    '36
    
    If ExisteArchivoFuente("INFTABELE", fecha - 2) Then
        LeerInfTabEleResultados fecha, 2
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 12 - No se encontro " & ArchivoInfTabEle(fecha - 2)
    End If
    '41
    
    If ExisteArchivoFuente("INFTABELE", fecha - 3) Then
        LeerInfTabEleResultados fecha, 3
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 13 - No se encontro " & ArchivoInfTabEle(fecha - 3)
    End If
    '46
    
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabEleDI fecha, 0
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 14 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '51
        
    If ExisteArchivoFuente("INFTABELE", fecha - 1) Then
        LeerInfTabEleDI fecha, 1
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 15 - No se encontro " & ArchivoInfTabEle(fecha - 1)
    End If
    '56
        
    If ExisteArchivoFuente("INFTABELE", fecha - 2) Then
        LeerInfTabEleDI fecha, 2
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 16 - No se encontro " & ArchivoInfTabEle(fecha - 2)
    End If
    '61
        
    If ExisteArchivoFuente("INFTABELE", fecha - 3) Then
        LeerInfTabEleDI fecha, 3
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 17 - No se encontro " & ArchivoInfTabEle(fecha - 3)
     End If
    '66
        
    If ExisteArchivoFuente("DDEC", fecha) Then
        LeerDDecGenProg fecha, 0
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM " 18 - No se encontro " & ArchivoDDEC(fecha)
    End If
    '69
    
    If ExisteArchivoFuente("DDEC", fecha - 1) Then
        LeerDDecGenProg fecha, 1
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM " 19 - No se encontro " & ArchivoDDEC(fecha - 1)
    End If
    '72
        
    If ExisteArchivoFuente("DDEC", fecha) Then
        LeerDDecGenProgHoraria fecha, 0
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM " 20 - No se encontro " & ArchivoDDEC(fecha)
    End If
    '75
        
    If ExisteArchivoFuente("IDO", fecha - 1) Then
        LeerIDOGenProg fecha, 1
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM " 21 - No se encontro " & ArchivoIDO(fecha - 1)
    End If
    '78
     
    If ExisteArchivoFuente("IDO", fecha - 1) Then
        LeerIDOGenRealEmpresa fecha, 1
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM " 22 - No se encontro " & ArchivoIDO(fecha - 1)
    End If
    '81
        
    If ExisteArchivoFuente("IDO", fecha - 2) Then
        LeerIDOGenRealEmpresa fecha, 2
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM " 23 - No se encontro " & ArchivoIDO(fecha - 2)
    End If
    '84
     
    If ExisteArchivoFuente("IDO", fecha - 3) Then
        LeerIDOGenRealEmpresa fecha, 3
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM " 24 - No se encontro " & ArchivoIDO(fecha - 3)
    End If
    '87
    
    If ExisteArchivoFuente("DAGC", fecha) Then
        LeerDAGC fecha
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM " 25 - No se encontro " & ArchivoDAGC(fecha)
    End If
    '90
    
    LogOfertaEPM "------SENSIBILIDADES---------", 3
    If ExisteArchivoFuente("INFSEN", fecha + 1) Then
        LeerArchivoSensibilidadesMobe fecha + 1
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM " 26 - No se encontro " & ArchivoInfSen(fecha + 1)
    End If
    
    LogOfertaEPM "------BALANCE---------", 3
    If ExisteArchivoFuente("IDO", fecha - 1) Then
        LeerIDOEmbalses fecha, 1
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM " 27 - No se encontro " & ArchivoIDO(fecha - 1)
    End If
    '92
    
    If ExisteArchivoFuente("IDO", fecha - 1) Then
        LeerIDORios fecha, 1
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM " 28 - No se encontro " & ArchivoIDO(fecha - 1)
    End If
    '94
    
    If ExisteArchivoFuente("DDEC", fecha) Then
        LeerDDecGeneracion fecha
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM " 29 - No se encontro " & ArchivoDDEC(fecha)
    End If
    '96
    
    If ExisteArchivoFuente("OFEI", fecha) Then
        LeerSEGDES fecha
        Avance = Avance + 1
        LeerOFEI fecha
        Avance = Avance + 1
        LeerOfeiDispPlanta fecha
        Avance = Avance + 1
        MostrarAvance Avance
    Else
        LogOfertaEPM " 30 - No se encontro " & ArchivoOFEI(fecha)
    End If
    '99
        
    DibujarEmbalses
    Avance = Avance + 1
    MostrarAvance Avance
    
    TiempoFinal = Time
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "OfertaEPM"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(fecha) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "OfertaEPM"
    End If

    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Number & " " & Err.Description
    Resume Next
End Sub



Sub InformeTableros(fecha As Date)
    Dim strArchivo As String
    Dim mensaje As String
    
    NroErrores = 0
 
    On Error GoTo ManejadorError
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    ThisWorkbook.Worksheets("Parametros").Cells(1, 4).Value = Usuario()
    ThisWorkbook.Worksheets("Parametros").Cells(1, 5).Value = "Fecha: " & CStr(Date) & " " & CStr(Time)
    TiempoInicio = Time
    
    Avance = 0
    NroErrores = 0
    blnUsarRutaAlterna = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamUsarRutaAlterna).Value
    blnCopiarEnRutaAlterna = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamCopiarEnRutaAlterna, ColParamUsarRutaAlterna).Value
    If blnCopiarEnRutaAlterna Then CopiarEnRutaAlterna fecha
    
    LogOfertaEPM "Inicio OfertaEPM Versión " & VersionOfertaEPM & " para la fecha " & CStr(fecha) & ", Ejecutado el dia: " & CStr(Date) & "  " & CStr(Time()) & "  " & Usuario(), 3
    LogOfertaEPM "TABLEROS", 3
    
    If Not (blnUsarRutaAlterna) Then CopiarArchivoIDOdesdeDiario fecha - 1
    
    If ExisteArchivoFuente("PRID", fecha) Then
        LeerTxtPreIdeal fecha
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM "  1 - No se encontro " & ArchivoPreideal(fecha)
    End If
    '3
    
    If ExisteArchivoFuente("DDEC", fecha) Then
        LeerDDEC fecha
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM "  2 - No se encontro " & ArchivoDDEC(fecha)
    End If
    '6
    
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabElePreciosOferta fecha, Prioridades, 0
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM "  3 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '8
        
    If ExisteArchivoFuente("INFTABELE", fecha - 1) Then
        LeerInfTabElePreciosOferta fecha, Prioridades, 1
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM "  4 - No se encontro " & ArchivoInfTabEle(fecha - 1)
    End If
    '10
        
    If ExisteArchivoFuente("INFTABELE", fecha - 2) Then
        LeerInfTabElePreciosOferta fecha, Prioridades, 2
        Avance = Avance + 2
        MostrarAvance Avance
    Else
        LogOfertaEPM "  5 - No se encontro " & ArchivoInfTabEle(fecha - 2)
    End If
    '12
        
    If ExisteArchivoFuente("INFTABELE", fecha - 3) Then
        LeerInfTabElePreciosOferta fecha, Prioridades, 3
        Avance = Avance + 4
        MostrarAvance Avance
    Else
        LogOfertaEPM "  6 - No se encontro " & ArchivoInfTabEle(fecha - 3)
    End If
    '16
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabElePrioridades fecha
        Avance = Avance + 4
        MostrarAvance Avance
    Else
        LogOfertaEPM "  7 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '20
        
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabElePrecioBolsa fecha
        Avance = Avance + 4
        MostrarAvance Avance
    Else
        LogOfertaEPM "  8 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '24
      
    If ExisteArchivoFuente("DMAR", fecha) Then
        LeerDMAR fecha
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM "  9 - No se encontro " & ArchivoDMAR(fecha)
    End If
    '27
    
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabEleResultados fecha, 0
        Avance = Avance + 4
        MostrarAvance Avance
    Else
        LogOfertaEPM " 10 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '31

    If ExisteArchivoFuente("INFTABELE", fecha - 1) Then
        LeerInfTabEleResultados fecha, 1
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 11 - No se encontro " & ArchivoInfTabEle(fecha - 1)
    End If
    '36
    
    If ExisteArchivoFuente("INFTABELE", fecha - 2) Then
        LeerInfTabEleResultados fecha, 2
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 12 - No se encontro " & ArchivoInfTabEle(fecha - 2)
    End If
    '41
    
    If ExisteArchivoFuente("INFTABELE", fecha - 3) Then
        LeerInfTabEleResultados fecha, 3
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 13 - No se encontro " & ArchivoInfTabEle(fecha - 3)
    End If
    '46
    
    If ExisteArchivoFuente("INFTABELE", fecha) Then
        LeerInfTabEleDI fecha, 0
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 14 - No se encontro " & ArchivoInfTabEle(fecha)
    End If
    '51
        
    If ExisteArchivoFuente("INFTABELE", fecha - 1) Then
        LeerInfTabEleDI fecha, 1
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 15 - No se encontro " & ArchivoInfTabEle(fecha - 1)
    End If
    '56
        
    If ExisteArchivoFuente("INFTABELE", fecha - 2) Then
        LeerInfTabEleDI fecha, 2
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 16 - No se encontro " & ArchivoInfTabEle(fecha - 2)
    End If
    '61
        
    If ExisteArchivoFuente("INFTABELE", fecha - 3) Then
        LeerInfTabEleDI fecha, 3
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 17 - No se encontro " & ArchivoInfTabEle(fecha - 3)
     End If
    '66
        
    If ExisteArchivoFuente("DDEC", fecha) Then
        LeerDDecGenProg fecha, 0
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM " 18 - No se encontro " & ArchivoDDEC(fecha)
    End If
    '69
    
    If ExisteArchivoFuente("DDEC", fecha - 1) Then
        LeerDDecGenProg fecha, 1
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM " 19 - No se encontro " & ArchivoDDEC(fecha - 1)
    End If
    '72
        
    If ExisteArchivoFuente("DDEC", fecha) Then
        LeerDDecGenProgHoraria fecha, 0
        Avance = Avance + 3
        MostrarAvance Avance
    Else
        LogOfertaEPM " 20 - No se encontro " & ArchivoDDEC(fecha)
    End If
    '75
        
    If ExisteArchivoFuente("IDO", fecha - 1) Then
        LeerIDOGenProg fecha, 1
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 21 - No se encontro " & ArchivoIDO(fecha - 1)
    End If
    '80
     
    If ExisteArchivoFuente("IDO", fecha - 1) Then
        LeerIDOGenRealEmpresa fecha, 1
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 22 - No se encontro " & ArchivoIDO(fecha - 1)
    End If
    '85
        
    If ExisteArchivoFuente("IDO", fecha - 2) Then
        LeerIDOGenRealEmpresa fecha, 2
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 23 - No se encontro " & ArchivoIDO(fecha - 2)
    End If
    '90
     
    If ExisteArchivoFuente("IDO", fecha - 3) Then
        LeerIDOGenRealEmpresa fecha, 3
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 24 - No se encontro " & ArchivoIDO(fecha - 3)
    End If
    '95
    
    If ExisteArchivoFuente("DAGC", fecha) Then
        LeerDAGC fecha
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM " 25 - No se encontro " & ArchivoDAGC(fecha)
    End If
    '100
    
    MostrarAvance Avance
    
    TiempoFinal = Time
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "Tableros"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(fecha) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "Tableros"
    End If

    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Number & " " & Err.Description
    Resume Next
End Sub

Sub InformeBalance(fecha As Date)
    Dim mensaje As String
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Avance = 0
    NroErrores = 0
   
    On Error GoTo ManejadorError
    Avance = 0
    MostrarAvance Avance
   
    LogOfertaEPM "", 2
    LogOfertaEPM "BALANCE", 3
    TiempoInicio = Time
    If ExisteArchivoFuente("IDO", fecha - 1) Then
        LeerIDOEmbalses fecha, 1
        Avance = Avance + 20
        MostrarAvance Avance
    Else
        LogOfertaEPM "No se encontro archivo IDO de la fecha " & CStr(fecha - 1) & Chr(10)
    End If
        
    If ExisteArchivoFuente("IDO", fecha - 1) Then
        LeerIDORios fecha, 1
        Avance = Avance + 20
        MostrarAvance Avance
    Else
        LogOfertaEPM "No se encontro archivo IDO de la fecha " & CStr(fecha - 1) & Chr(10)
    End If
        
    If ExisteArchivoFuente("DDEC", fecha) Then
        LeerDDecGeneracion fecha
        Avance = Avance + 20
        MostrarAvance Avance
    Else
        LogOfertaEPM "No se encontro archivo DDEC de la fecha " & CStr(fecha) & Chr(10)
    End If
        
    If ExisteArchivoFuente("OFEI", fecha) Then
        LeerSEGDES fecha
        LeerOFEI fecha
        LeerOfeiDispPlanta fecha
        Avance = Avance + 20
        MostrarAvance Avance
    Else
        LogOfertaEPM "No se encontro archivo OFEI de la fecha " & CStr(fecha) & Chr(10)
    End If
        
    DibujarEmbalses
    Avance = Avance + 20
    MostrarAvance Avance
    
    TiempoFinal = Time
    
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "Balance"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(fecha) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "Balance"
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Exit Sub
ManejadorError:
    strError = Err.Description & Chr(10) & strError & Chr(10)
    LogOfertaEPM strError
    Resume Next
End Sub


Sub RevisarPrecioMarginal(fecha As Date)
    Dim mensaje As String
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    strError = ""
    NroErrores = 0
    Avance = 0
    MostrarAvance 0
    On Error GoTo ManejadorError
    
    Avance = 0
    MostrarAvance Avance
    TiempoInicio = Time
    LogOfertaEPM "", 2
    LogOfertaEPM "REVISAR PRECIO MARGINAL", 3
    If ExisteArchivoFuente("PRID", fecha) Then
        
        LeerTxtPreIdeal fecha
        Avance = Avance + 25
        MostrarAvance Avance
    Else
        LogOfertaEPM "No se encontro archivo PreIdeal de la fecha " & CStr(fecha)
    End If
    '25
    
    If ExisteArchivoFuente("DSEGDES", fecha) Then
        LeerSEGDES fecha
        Avance = Avance + 25
        MostrarAvance Avance
    Else
        LogOfertaEPM "No se encontro archivo DSEGDES de la fecha " & CStr(fecha)
    End If
    '50
    
    If ExisteArchivoFuente("DDEC", fecha) Then
        LeerDDEC fecha
        Avance = Avance + 10
        MostrarAvance Avance
        LeerDDecGenProg fecha, 0
        Avance = Avance + 5
        MostrarAvance Avance
        LeerDDecGenProgHoraria fecha, 0
        Avance = Avance + 5
        MostrarAvance Avance
    Else
        LogOfertaEPM "No se encontro archivo DDEC de la fecha " & CStr(fecha) & Chr(10)
        NroErrores = NroErrores + 1
    End If
    '70
    
    If ExisteArchivoFuente("DDEC", fecha - 1) Then
        LeerDDecGenProg fecha, 1
        Avance = Avance + 10
        MostrarAvance Avance
    Else
        LogOfertaEPM "No se encontro archivo DDEC de la fecha " & CStr(fecha - 1) & Chr(10)
        NroErrores = NroErrores + 1
    End If
    '80
   
    If ExisteArchivoFuente("OFEI", fecha) Then
        LeerOFEI fecha
        Avance = Avance + 20
        MostrarAvance Avance
    Else
        LogOfertaEPM "No se encontro archivo OFEI de la fecha " & CStr(fecha) & Chr(10)
        NroErrores = NroErrores + 1
    End If
    '100
    TiempoFinal = Time
    
    If NroErrores > 0 Then
        mensaje = "Finalizo con " & CStr(NroErrores) & " errores" & Chr(10) & "en " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        Beep
        MsgBox mensaje, vbOKOnly, "RevisarPrecioMarginal"
        Shell "Notepad.exe " & strArchivoLog, vbNormalFocus
    Else
        mensaje = "Fin OfertaEPM Versión  " & VersionOfertaEPM & " " & CStr(fecha) & Chr(10) & "En " & CStr(Round((TiempoFinal - TiempoInicio) * 86400, 0)) & " s"
        LogOfertaEPM mensaje, 3
        MsgBox mensaje, vbOKOnly, "RevisarPrecioMarginal"
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Exit Sub
ManejadorError:
    NroErrores = NroErrores + 1
    strError = Err.Description & Chr(10) & strError & Chr(10)
    LogOfertaEPM strError
    Resume Next
End Sub


