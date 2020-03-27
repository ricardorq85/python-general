Option Explicit

Public blnUsarRutaAlterna As Boolean
Public blnCopiarEnRutaAlterna As Boolean

Public FechaReal As Date
Public casoOfertado As Integer

Public Const FilaParamFechaReal = 1
Public Const FilaParamOFEI = 3
Public Const FilaParamDDEC = 4
Public Const FilaParamDAGC = 5
Public Const FilaParamDMAR = 6
Public Const FilaParamSEGDES = 7
Public Const FilaParamPrid = 8
Public Const FilaParamIDO = 9
Public Const FilaParamInfTabEle = 10
Public Const FilaParamInfSen = 11
Public Const FilaParamIMAR = 12
Public Const FilaParamDIARIO = 13
Public Const FilaParamRutaAlterna = 14
Public Const FilaParamCopiarEnRutaAlterna = 15

Public Const FilaParamIDOPalabras = 18

Public Const ColParamFechaReal = 2
Public Const ColParamRaiz = 2
Public Const ColParamPrefijo = 1
Public Const ColParamCodPalabra = 1
Public Const ColParamPalabraIDO = 2
Public Const ColParamFilaIDO = 3
Public Const ColParamUsarRutaAlterna = 4

Public Const ColEmbEmbalseIDO = 1
Public Const ColEmbVolUtil = 2
Public Const ColEmbVolInicialPorc = 3
Public Const ColEmbVolInicialMm3 = 4
Public Const ColEmbPosX = 5
Public Const ColEmbPosY = 6
Public Const ColEmbTamX = 7
Public Const ColEmbTamY = 8
Public Const ColEmbCoefA = 11
Public Const ColEmbCoefB = 12
Public Const ColEmbCoefC = 13
Public Const ColEmbCoefD = 14
Public Const ColEmbCentral = 16

Public Const ColRiosEmbalse = 1
Public Const ColRiosRioIDO = 2
Public Const ColRiosCaudal = 3
Public Const ColRiosPorc = 5
Public Const ColRiosGWhDia = 6

Public Const ColSenCasoOfertado = 3
Public Const FilaSenCasoOfertado = 1

Public Const MAXEMPRESAS = 20
Public Const MAXCENTRALES = 500
Public Const MAXEMBALSES = 50
Public Const MAXRIOS = 50
Public Const MAXUNIDADES = 500

Public Enum nmMes
    largo = 1    'Mes largo: Enero, Febrero, ...
    Corto = 2     'Mes corto: Ene, Feb, ...
    NumeroConCero = 3      'mm    - 01, 02, 03
    NumeroSinCero = 4       'm     - 1,2,3
End Enum

Public Enum nmDia
    largo = 1    'Dia largo: Lunes, Martes, ...
    Corto = 2    'Dia corto: lun, mar, ...
End Enum




Public Enum nroDia
    ConCero = 1    '01, 02, ... , 31
    SinCero = 2    '1, 2, ... , 31
End Enum



Public Function NombreMes(Tipo As nmMes, fecha As Date) As String
    'MMMM  - Mes largo: Enero, Febrero, ...
    'mmm   - Mes corto: Ene, Feb, ...
    'mm    - 01, 02, 03
    'm     - 1,2,3
    Dim nroMes As Integer
    Dim mesAux As String
    nroMes = Month(fecha)
    
    If Tipo = nmMes.largo Then

            If nroMes = 1 Then mesAux = "Enero"
            If nroMes = 2 Then mesAux = "Febrero"
            If nroMes = 3 Then mesAux = "Marzo"
            If nroMes = 4 Then mesAux = "Abril"
            If nroMes = 5 Then mesAux = "Mayo"
            If nroMes = 6 Then mesAux = "Junio"
            If nroMes = 7 Then mesAux = "Julio"
            If nroMes = 8 Then mesAux = "Agosto"
            If nroMes = 9 Then mesAux = "Septiembre"
            If nroMes = 10 Then mesAux = "Octubre"
            If nroMes = 11 Then mesAux = "Noviembre"
            If nroMes = 12 Then mesAux = "Diciembre"
    End If
    If Tipo = nmMes.Corto Then

            If nroMes = 1 Then mesAux = "Ene"
            If nroMes = 2 Then mesAux = "Feb"
            If nroMes = 3 Then mesAux = "Mar"
            If nroMes = 4 Then mesAux = "Abr"
            If nroMes = 5 Then mesAux = "May"
            If nroMes = 6 Then mesAux = "Jun"
            If nroMes = 7 Then mesAux = "Jul"
            If nroMes = 8 Then mesAux = "Ago"
            If nroMes = 9 Then mesAux = "Sep"
            If nroMes = 10 Then mesAux = "Oct"
            If nroMes = 11 Then mesAux = "Nov"
            If nroMes = 12 Then mesAux = "Dic"
    End If
    
    If Tipo = NumeroConCero Then

            If nroMes = 1 Then mesAux = "01"
            If nroMes = 2 Then mesAux = "02"
            If nroMes = 3 Then mesAux = "03"
            If nroMes = 4 Then mesAux = "04"
            If nroMes = 5 Then mesAux = "05"
            If nroMes = 6 Then mesAux = "06"
            If nroMes = 7 Then mesAux = "07"
            If nroMes = 8 Then mesAux = "08"
            If nroMes = 9 Then mesAux = "09"
            If nroMes = 10 Then mesAux = "10"
            If nroMes = 11 Then mesAux = "11"
            If nroMes = 12 Then mesAux = "12"
    End If
    
    If Tipo = NumeroSinCero Then

            If nroMes = 1 Then mesAux = "1"
            If nroMes = 2 Then mesAux = "2"
            If nroMes = 3 Then mesAux = "3"
            If nroMes = 4 Then mesAux = "4"
            If nroMes = 5 Then mesAux = "5"
            If nroMes = 6 Then mesAux = "6"
            If nroMes = 7 Then mesAux = "7"
            If nroMes = 8 Then mesAux = "8"
            If nroMes = 9 Then mesAux = "9"
            If nroMes = 10 Then mesAux = "10"
            If nroMes = 11 Then mesAux = "11"
            If nroMes = 12 Then mesAux = "12"
    End If
        
    NombreMes = mesAux
End Function

Public Function nroDia(Tipo As nroDia, fecha) As String
    Dim dia As Integer
    Dim strDia As String
    dia = Day(fecha)
    strDia = CStr(dia)
    If Len(strDia) = 1 And Tipo = ConCero Then strDia = "0" & strDia
    nroDia = strDia
End Function

Public Function NombreDia(Tipo As nmDia, fecha) As String
    Dim dia As Integer
    Dim strDia As String
    dia = Weekday(fecha, vbMonday)
    Select Case dia
        Case 1
            If Tipo = nmDia.Corto Then strDia = "lun" Else strDia = "Lunes"
        Case 2
            If Tipo = nmDia.Corto Then strDia = "mar" Else strDia = "Martes"
        Case 3
            If Tipo = nmDia.Corto Then strDia = "mie" Else strDia = "Miercoles"
        Case 4
            If Tipo = nmDia.Corto Then strDia = "jue" Else strDia = "Jueves"
        Case 5
            If Tipo = nmDia.Corto Then strDia = "vie" Else strDia = "Viernes"
        Case 6
            If Tipo = nmDia.Corto Then strDia = "sab" Else strDia = "Sabado"
        Case 7
            If Tipo = nmDia.Corto Then strDia = "dom" Else strDia = "Domingo"
    End Select
    
    NombreDia = strDia
End Function



Public Function DejarUnSoloEspacio(Linea As String) As String

    Dim LineaSalida As String
    Dim letra As String
    Dim Longitud As Integer
    Dim i As Integer
    Dim blnPrimerEspacio As Boolean
    
    Longitud = Len(Linea)
    blnPrimerEspacio = False
    For i = 1 To Longitud
        letra = Mid(Linea, i, 1)
        If letra <> Chr(32) And letra <> Chr(160) Then
            LineaSalida = LineaSalida & letra
            blnPrimerEspacio = False
        Else
            If blnPrimerEspacio = False Then
                LineaSalida = LineaSalida & letra
                blnPrimerEspacio = True
            End If
        End If
    Next i
    If Len(LineaSalida) > 1 Then
        DejarUnSoloEspacio = Trim(LineaSalida)
    Else
        DejarUnSoloEspacio = LineaSalida
    End If
End Function

Public Function EliminarComillas(Cadena As String) As String

    Dim cadenaSalida As String
    Dim letra As String
    Dim Longitud As Integer
    Dim i As Integer
    
    Longitud = Len(Cadena)

    For i = 1 To Longitud
        letra = Mid(Cadena, i, 1)
        If letra <> Chr(34) Then
            cadenaSalida = cadenaSalida & letra
        End If
    Next i

    EliminarComillas = Trim(cadenaSalida)

End Function


Public Function CodigoHash(nm As String) As Long

    Dim i As Integer
    Dim Valor As Integer
    Dim Suma As Integer
    Dim largo As Integer
    Dim letra As String
    Dim Inicio As String
    Dim Final As String
    largo = Len(nm)
    Suma = 0
    letra = ""
    
    For i = 1 To largo
        letra = Mid(nm, i, 1)
        If i = 1 Then Valor = Asc(letra) * 2
        If i = 2 Then Valor = Asc(letra) * 3
        If i = 3 Then Valor = Asc(letra) * 5
        If i = 4 Then Valor = Asc(letra) * 7
        If i = 5 Then Valor = Asc(letra) * 11
        If i = 6 Then Valor = Asc(letra) * 13
        If i = 7 Then Valor = Asc(letra) * 17
        If i = 8 Then Valor = Asc(letra) * 19
        If i = 9 Then Valor = Asc(letra) * 21
        If i = 10 Then Valor = Asc(letra) * 23
        If i = 11 Then Valor = Asc(letra) * 29
        If i > 11 Then Valor = Asc(letra) * 31
        Suma = Suma + Valor
    Next i
    
    If nm = "MCARUQUIA" Then Suma = Suma + 5
    
    CodigoHash = Suma

End Function

Public Function CodigoHashCorto(nm As String) As Long

    Dim i As Integer
    Dim Valor As Integer
    Dim Suma As Integer
    Dim largo As Integer
    Dim letra As String
    Dim Inicio As String
    Dim Final As String
    largo = Len(nm)
    If largo > 5 Then largo = 5
    Suma = 0
    letra = ""
    
    For i = 1 To largo
        letra = Mid(nm, i, 1)
        Valor = Asc(letra)
        If Valor > 60 Then Valor = Valor - 60
        If i = 1 Then Valor = Valor * 2
        If i = 2 Then Valor = Valor * 3
        If i = 3 Then Valor = Valor * 5
        If i = 4 Then Valor = Valor * 7
        If i = 5 Then Valor = Valor * 11


        Suma = Suma + Valor
    Next i
    
    CodigoHashCorto = Suma

End Function

Public Function FCEmb(prmEmb As String, porc As Double) As Double
    
    Dim fila As Integer
    Dim nmEmb As String
    Dim CoefA As Double
    Dim CoefB As Double
    Dim CoefC As Double
    Dim CoefD As Double
    Dim blnHallado As Boolean
    Dim strEmb As String
    
    strEmb = UCase(Trim(prmEmb))
    
    fila = 3
    blnHallado = False
    nmEmb = UCase(Trim(ThisWorkbook.Worksheets("Embalses").Cells(fila, ColEmbEmbalseIDO).Value))
    Do While nmEmb <> "" And blnHallado = False
        DoEvents
        

        If nmEmb = strEmb Then
            CoefA = ThisWorkbook.Worksheets("Embalses").Cells(fila, ColEmbCoefA).Value
            CoefB = ThisWorkbook.Worksheets("Embalses").Cells(fila, ColEmbCoefB).Value
            CoefC = ThisWorkbook.Worksheets("Embalses").Cells(fila, ColEmbCoefC).Value
            CoefD = ThisWorkbook.Worksheets("Embalses").Cells(fila, ColEmbCoefD).Value
            
            blnHallado = True
        
        End If
        fila = fila + 1
        nmEmb = UCase(Trim(ThisWorkbook.Worksheets("Embalses").Cells(fila, ColEmbEmbalseIDO).Value))
    Loop
    FCEmb = CoefA * porc ^ 3 + CoefB * porc ^ 2 + CoefC * porc + CoefD
End Function

'Consulta en la hoja rios los aportes de los rios y devuelve el promedio que es utilizado en la hoja Embalses
'Calcula el porcentaje ponderado de aportes cuando varios rios aportan a un embalse
'Si m3s = true AportesEmbalse devuelve el valor en m3/s
Public Function AportesEmbalse(nmEmb As String, Optional m3s As Boolean = False) As Double
    Dim fila As Integer
    Dim Valor As Single
    Dim Suma As Single
    Dim strEmb As String
    Dim embalse As String
    Dim Caudal As Single
    Dim caudalTotal As Single
    Dim PorcRio As Single
    Dim PorcTotal As Single
    Dim PromRio As Single
    Dim PromTotal As Single
    
    fila = 2
    Valor = 0
    Suma = 0
    Caudal = 0
    PorcRio = 0
    PorcTotal = 0
    PromRio = 0
    PromTotal = 0
    strEmb = UCase(Trim(nmEmb))
    AportesEmbalse = 0
    embalse = UCase(Trim(ThisWorkbook.Worksheets("Rios").Cells(fila, ColRiosEmbalse).Value))
    Do While embalse <> "FINAL"
        DoEvents
        If embalse = strEmb Then
            'Acumular caudal
            Caudal = ThisWorkbook.Worksheets("Rios").Cells(fila, ColRiosCaudal).Value
            caudalTotal = caudalTotal + Caudal
            
            PorcRio = ThisWorkbook.Worksheets("Rios").Cells(fila, ColRiosPorc).Value
            If PorcRio <> 0 Then
                PromRio = Caudal * 100 / PorcRio
                PromTotal = PromTotal + PromRio
                PorcTotal = caudalTotal / PromTotal
            End If
            
            AportesEmbalse = PorcTotal * 100
        End If
        fila = fila + 1
        embalse = UCase(Trim(ThisWorkbook.Worksheets("Rios").Cells(fila, ColRiosEmbalse).Value))
    Loop
    
    If m3s = True Then AportesEmbalse = caudalTotal
    
End Function

Sub FormatoSimpleHoja(nmHoja As String)
    Dim nroFilas As Integer
    nroFilas = ThisWorkbook.Worksheets(nmHoja).UsedRange.Rows.Count
    With ThisWorkbook.Worksheets(nmHoja).UsedRange.Interior
        '.Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With ThisWorkbook.Worksheets(nmHoja).Range(Cells(1, 1).Address, Cells(nroFilas, 1).Address).Interior
        '.Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With ThisWorkbook.Worksheets(nmHoja).Range(Cells(1, 1).Address, Cells(1, 26).Address).Interior
        '.Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
End Sub





Public Function CentralEsMayor(nmCen As String) As Boolean
    Dim nmCentral As String
    Dim fila As Integer
    Dim tipoCen As String
    Dim blnHallada As Boolean
    
    fila = 2
    CentralEsMayor = False
    blnHallada = False
    
    nmCentral = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(fila, colEquivCentralDDEC).Value))
    Do While blnHallada = False And fila < MAXCENTRALES
        DoEvents
        If nmCentral = nmCen Then
            tipoCen = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(fila, colEquivTipo).Value))
            If tipoCen = "GH" Or tipoCen = "GT" Then
                CentralEsMayor = True
            End If
            blnHallada = True
            'Debug.Print Fila
        End If
    
        fila = fila + 1
        nmCentral = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(fila, colEquivCentralDDEC).Value))
    Loop
End Function

Function MIN(a As Single, b As Single) As Single
    If a > b Then MIN = b Else MIN = a
End Function

Public Function ExtensionExcel(xlsArchivo As String) As String
    Dim nmArchivoXLSX As String
    Dim nmArchivoXLS As String
    Dim nmArchivoXLSM As String
    Dim nmArchivoXLSB As String
    
    nmArchivoXLSX = UCase(Trim(xlsArchivo)) & ".XLSX"
    nmArchivoXLS = UCase(Trim(xlsArchivo)) & ".XLS"
    nmArchivoXLSM = UCase(Trim(xlsArchivo)) & ".XLSM"
    nmArchivoXLSB = UCase(Trim(xlsArchivo)) & ".XLSB"
    
    ExtensionExcel = ""
    If Len(Dir(nmArchivoXLSX)) > 0 Then
        ExtensionExcel = ".XLSX"
    
    ElseIf Len(Dir(nmArchivoXLS)) > 0 Then
        ExtensionExcel = ".XLS"
    
    ElseIf Len(Dir(nmArchivoXLSM)) > 0 Then
        ExtensionExcel = ".XLSM"
    
    ElseIf Len(Dir(nmArchivoXLSB)) > 0 Then
        ExtensionExcel = ".XLSB"
    
    End If
End Function


Public Function ExisteArchivoFuente(Tipo As String, fecha As Date) As Boolean
    Dim strTipo As String
    Dim strNombreArchivo As String
    
    ExisteArchivoFuente = False
    strTipo = UCase(Trim(Tipo))
    If strTipo = "OFEI" Then strNombreArchivo = ArchivoOFEI(fecha)
    If strTipo = "DDEC" Then strNombreArchivo = ArchivoDDEC(fecha)
    If strTipo = "DAGC" Then strNombreArchivo = ArchivoDAGC(fecha)
    If strTipo = "DMAR" Then strNombreArchivo = ArchivoDDEC(fecha)
    If strTipo = "DSEGDES" Then strNombreArchivo = ArchivoDSEGDES(fecha)
    If strTipo = "PRID" Then strNombreArchivo = ArchivoPreideal(fecha)
    If strTipo = "IDO" Then strNombreArchivo = ArchivoIDO(fecha)
    If strTipo = "INFTABELE" Then strNombreArchivo = ArchivoInfTabEle(fecha)
    If strTipo = "INFSEN" Then strNombreArchivo = ArchivoInfSen(fecha)
    If strTipo = "IMAR" Then strNombreArchivo = ArchivoIMAR(fecha)
    
    If Len(Dir(strNombreArchivo)) > 0 Then
        ExisteArchivoFuente = True
    End If
    
End Function





Public Sub RevisarFuentes(fecha As Date)
    Dim strArchivoOFEI As String
    Dim strArchivodDEC As String
    Dim strArchivodAGC As String
    Dim strArchivodMAR As String
    Dim strArchivodSEGDES As String
    Dim strArchivoPrid As String
    Dim strArchivoIDO As String
    Dim strArchivoInfTabEle As String
    Dim strArchivoInfSen As String
    Dim strArchivoIMAR As String
    
    strArchivoOFEI = ArchivoOFEI(fecha)
    strArchivodDEC = ArchivoDDEC(fecha)
    strArchivodAGC = ArchivoDAGC(fecha)
    strArchivodMAR = ArchivoDMAR(fecha)
    strArchivodSEGDES = ArchivoDSEGDES(fecha)
    strArchivoPrid = ArchivoPreideal(fecha)
    strArchivoIDO = ArchivoIDO(fecha)
    strArchivoInfTabEle = ArchivoInfTabEle(fecha)
    strArchivoInfSen = ArchivoInfSen(fecha)
    strArchivoIMAR = ArchivoIMAR(fecha)
    
    If Len(Dir(strArchivoOFEI)) > 0 Then
        Debug.Print "OFEI OK"
    End If
    If Len(Dir(strArchivodDEC)) > 0 Then
        Debug.Print "DDEC OK"
    End If
    If Len(Dir(strArchivodAGC)) > 0 Then
        Debug.Print "DAGC OK"
    End If
    If Len(Dir(strArchivodMAR)) > 0 Then
        Debug.Print "DMAR OK"
    End If
    If Len(Dir(strArchivodSEGDES)) > 0 Then
        Debug.Print "SegDes OK"
    End If
    If Len(Dir(strArchivoPrid)) > 0 Then
        Debug.Print "Preideal OK"
    End If
    If Len(Dir(strArchivoIDO)) > 0 Then
        Debug.Print "IDO OK"
    End If
    If Len(Dir(strArchivoInfTabEle)) > 0 Then
        Debug.Print "InfTabEle OK"
    End If
    If Len(Dir(strArchivoInfSen)) > 0 Then
        Debug.Print "InfSen OK"
    End If
    If Len(Dir(strArchivoIMAR)) > 0 Then
        Debug.Print "IMAR OK"
    End If
    
End Sub

'Si incrementar = 1 se trata de un error
'Si incrementar = 2 registrar datos usuario, libro y fecha
'Si incrementar = 2 solo escribir el texto enviado
Public Sub LogOfertaEPM(texto As String, Optional incrementar As Integer = 1)
    Dim strCadena As String
    
    strArchivoLog = ThisWorkbook.Path & "\LogOfertaEPM" & CStr(Year(Date)) & "-" & CStr(Month(Date)) & "-" & CStr(Day(Date)) & ".txt"
    
    If incrementar = 1 Then 'Error
        Open strArchivoLog For Append As #9
        NroErrores = NroErrores + 1
        Print #9, "Error: " & NroErrores & " " & texto
    End If
    
    If incrementar = 2 Then 'Registro
        Open strArchivoLog For Append As #9
        Print #9, "Usuario: " & Usuario() & ", Libro: " & ThisWorkbook.Name & ", Fecha: " & CStr(Date) & ", Hora: " & CStr(Time())
    End If

    If incrementar = 3 Then 'Solo mensaje
        Open strArchivoLog For Append As #9
        Print #9, texto
    End If
    
    Close #9

End Sub

Public Sub CopiarEnRutaAlterna(fecha As Date)

    Dim fso As Object
    Dim RutaDestino As String
    Dim strArchivo As String
    
    RutaDestino = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
    Set fso = CreateObject("Scripting.FileSystemObject")
 
        strArchivo = ArchivoOFEI(fecha)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
        
        strArchivo = ArchivoDDEC(fecha)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
        
        strArchivo = ArchivoDAGC(fecha)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
         
        strArchivo = ArchivoDMAR(fecha)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
        
        strArchivo = ArchivoDSEGDES(fecha)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
        
        strArchivo = ArchivoPreideal(fecha)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
        
        strArchivo = ArchivoIDO(fecha - 1)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
        
        strArchivo = ArchivoInfTabEle(fecha)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
            
        strArchivo = ArchivoInfSen(fecha)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
            
        strArchivo = ArchivoIMAR(fecha)
        If Len(Dir(strArchivo)) > 0 Then
            fso.copyfile strArchivo, RutaDestino
        Else
            LogOfertaEPM "CopiarEnRutaAlterna: No se encontro " & strArchivo
        End If
        
    Set fso = Nothing

End Sub

Function FormatoFechaWindows() As String
    Dim Separador As String
    With Application
        Separador = .International(xlDateSeparator)
        Select Case .International(xlDateOrder)
            Case Is = 0 '"mm dd yyyy"
                FormatoFechaWindows = "mm" & Separador & "dd" & Separador & "yyyy"
            Case Is = 1 '"dd mm yyyy"
                FormatoFechaWindows = "dd" & Separador & "mm" & Separador & "yyyy"
            Case Is = 2 '"yyyy mm dd"
                FormatoFechaWindows = "yyyy" & Separador & "mm" & Separador & "dd"
        End Select
    End With
End Function
