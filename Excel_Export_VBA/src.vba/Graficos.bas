Option Explicit

Public HojaGrafico As Worksheet
Public RefX As Integer     'Posicion central coordenada x = 0
Public RefY As Integer     'Posicion central coordenada y = 0
Public FactorX As Integer  'Factor de ampliacion
Public FactorY As Integer  'Factor de ampliacion

Public Type typeEmbalse
    id As Integer
    nombre As String
    porc As Double
    nep As Double
    central As String
    capCenGWhd As Single
    porcFD1 As Double
    porcFD2 As Double
    posX As Integer
    posY As Integer
    TamX As Integer
    TamY As Integer
    strForma As String
    objForma As shape
    objNivel As shape
    objNivelFinalD1 As shape
    objNivelFinalD2 As shape
    objNivelNEP As shape
    objPorcFD1 As shape
    objPorcFD2 As shape
    objValorNEP As shape
    objRiosPorc As shape
    objGenPorc As shape
End Type


Public Embalses(MAXEMBALSES) As typeEmbalse

Public Sub DibujarEmbalses()
    Dim nroEmb As Integer
    
    Dim nmEmbalse As String
    Dim nmCen As String
    Dim filaEmb As Integer

    On Error GoTo ManejadorError
    FechaReal = ThisWorkbook.Worksheets("Parametros").Cells(1, 2).Value
    ThisWorkbook.Worksheets("Grafico").Cells(1, 1).Value = "Fecha: " & CStr(FechaReal)
    
    If HojaGrafico Is Nothing Then Set HojaGrafico = ThisWorkbook.Worksheets("Grafico")
    

    HojaGrafico.Activate
    HojaGrafico.Shapes.SelectAll
    Selection.Delete
    
    nroEmb = 0
    filaEmb = 3
    nmEmbalse = UCase(Trim(ThisWorkbook.Worksheets("Embalses").Cells(filaEmb, ColEmbEmbalseIDO).Value))
    Do While nmEmbalse <> "" And nmEmbalse <> "TOTAL SIN"
        nroEmb = nroEmb + 1
        
        Embalses(nroEmb).nombre = nmEmbalse
        Embalses(nroEmb).id = nroEmb
        nmCen = ThisWorkbook.Worksheets("Embalses").Cells(filaEmb, ColEmbCentral).Value
        Embalses(nroEmb).central = nmCen
        Embalses(nroEmb).capCenGWhd = ThisWorkbook.Worksheets("Embalses").Cells(filaEmb, 18).Value
        
        Embalses(nroEmb).nep = nep(nmEmbalse, FechaReal)
        Embalses(nroEmb).porc = ThisWorkbook.Worksheets("Embalses").Cells(filaEmb, ColEmbVolInicialPorc).Value
        Embalses(nroEmb).porcFD1 = InfoBal(nmEmbalse, "Volumen Final", "Embalse", "%", "hoy")
        Embalses(nroEmb).porcFD2 = InfoBal(nmEmbalse, "Volumen Final", "Embalse", "%", "sig")
        
        Embalses(nroEmb).posX = ThisWorkbook.Worksheets("Embalses").Cells(filaEmb, ColEmbPosX).Value
        Embalses(nroEmb).posY = ThisWorkbook.Worksheets("Embalses").Cells(filaEmb, ColEmbPosY).Value
        Embalses(nroEmb).TamX = ThisWorkbook.Worksheets("Embalses").Cells(filaEmb, ColEmbTamX).Value
        Embalses(nroEmb).TamY = ThisWorkbook.Worksheets("Embalses").Cells(filaEmb, ColEmbTamY).Value
        
        'Mostrar embalse maximo
        Set Embalses(nroEmb).objForma = HojaGrafico.Shapes.AddShape(msoShapeRectangle, Embalses(nroEmb).posX, Embalses(nroEmb).posY + 200 - Embalses(nroEmb).TamX, Embalses(nroEmb).TamX, Embalses(nroEmb).TamX)
        Embalses(nroEmb).objForma.TextFrame.Characters.Text = Mid(nmEmbalse, 1, 11) & Chr(10) & CStr(Round(Embalses(nroEmb).porc, 1)) & " %"
        If nep(Embalses(nroEmb).nombre, FechaReal) * 100 > Embalses(nroEmb).porc Then
            Embalses(nroEmb).objForma.TextFrame.Characters.Font.Color = vbRed
        Else
            Embalses(nroEmb).objForma.TextFrame.Characters.Font.Color = vbBlack
        End If
        Embalses(nroEmb).objForma.TextFrame.Characters.Font.Bold = True
        Embalses(nroEmb).objForma.TextFrame.Characters.Font.Size = 11
        Embalses(nroEmb).objForma.Fill.Transparency = 0.95
        
        'Mostrar nivel inicial dia 1
        Set Embalses(nroEmb).objNivel = HojaGrafico.Shapes.AddShape(msoShapeRectangle, Embalses(nroEmb).posX, Embalses(nroEmb).posY + (200 - Embalses(nroEmb).TamY), Embalses(nroEmb).TamX, Embalses(nroEmb).TamY)
        Embalses(nroEmb).objNivel.Fill.Transparency = 0.85
        If nep(Embalses(nroEmb).nombre, FechaReal) * 100 > Embalses(nroEmb).porc Then
            Embalses(nroEmb).objNivel.Line.ForeColor.RGB = RGB(255, 0, 0)
            Embalses(nroEmb).objNivel.Line.Weight = 2
        End If
        
        
        'Mostrar linea nivel final del dia 1
        Set Embalses(nroEmb).objNivelFinalD1 = HojaGrafico.Shapes.AddConnector(msoConnectorStraight, _
                    Embalses(nroEmb).posX, _
                    Embalses(nroEmb).posY + 200 - Embalses(nroEmb).TamX * Embalses(nroEmb).porcFD1 / 100, _
                    Embalses(nroEmb).posX + Embalses(nroEmb).TamX * 0.48, _
                    Embalses(nroEmb).posY + 200 - Embalses(nroEmb).TamX * Embalses(nroEmb).porcFD1 / 100)
        Embalses(nroEmb).objNivelFinalD1.Line.DashStyle = msoLineSolid
        Embalses(nroEmb).objNivelFinalD1.Line.Weight = 1
        Embalses(nroEmb).objNivelFinalD1.Line.ForeColor.RGB = RGB(55, 55, 55)
        Embalses(nroEmb).objNivelFinalD1.Line.BeginArrowheadStyle = msoArrowheadOpen
        If nep(Embalses(nroEmb).nombre, FechaReal + 1) * 100 > Embalses(nroEmb).porcFD1 Then
            Embalses(nroEmb).objNivelFinalD1.Line.ForeColor.RGB = RGB(255, 0, 0)
            Embalses(nroEmb).objNivelFinalD1.Line.Weight = 2
        End If
        

        'Mostrar etiqueta nivel final del dia 1
        Set Embalses(nroEmb).objPorcFD1 = HojaGrafico.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                    Embalses(nroEmb).posX - 60, _
                    Embalses(nroEmb).posY + 200 - Embalses(nroEmb).TamX * Embalses(nroEmb).porcFD1 / 100, _
                    60, 30)
        Embalses(nroEmb).objPorcFD1.TextFrame.Characters.Text = CStr(Round(Embalses(nroEmb).porcFD1, 1)) & " %"
        If nep(Embalses(nroEmb).nombre, FechaReal + 1) * 100 > Embalses(nroEmb).porcFD1 Then
            Embalses(nroEmb).objPorcFD1.TextFrame.Characters.Font.Color = vbRed
        Else
            Embalses(nroEmb).objPorcFD1.TextFrame.Characters.Font.Color = vbBlack
        End If
        Embalses(nroEmb).objPorcFD1.TextFrame.Characters.Font.Bold = True
        Embalses(nroEmb).objPorcFD1.Fill.Transparency = 0.9
        Embalses(nroEmb).objPorcFD1.Line.Visible = msoFalse
        Embalses(nroEmb).objPorcFD1.TextFrame.AutoSize = True
        
        'Mostrar linea nivel final del dia 2
        Set Embalses(nroEmb).objNivelFinalD2 = HojaGrafico.Shapes.AddConnector(msoConnectorStraight, _
                    Embalses(nroEmb).posX + Embalses(nroEmb).TamX * 0.52, _
                    Embalses(nroEmb).posY + 200 - Embalses(nroEmb).TamX * Embalses(nroEmb).porcFD2 / 100, _
                    Embalses(nroEmb).posX + Embalses(nroEmb).TamX, _
                    Embalses(nroEmb).posY + 200 - Embalses(nroEmb).TamX * Embalses(nroEmb).porcFD2 / 100)
        Embalses(nroEmb).objNivelFinalD2.Line.DashStyle = msoLineDash
        Embalses(nroEmb).objNivelFinalD2.Line.Weight = 1
        Embalses(nroEmb).objNivelFinalD2.Line.ForeColor.RGB = RGB(50, 50, 50)
        Embalses(nroEmb).objNivelFinalD2.Line.EndArrowheadStyle = msoArrowheadOpen
        If nep(Embalses(nroEmb).nombre, FechaReal + 2) * 100 > Embalses(nroEmb).porcFD2 Then
            Embalses(nroEmb).objNivelFinalD2.Line.ForeColor.RGB = RGB(255, 0, 0)
            Embalses(nroEmb).objNivelFinalD2.Line.Weight = 2
        End If

        'Mostrar etiqueta nivel final del dia 2
        Set Embalses(nroEmb).objPorcFD2 = HojaGrafico.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                    Embalses(nroEmb).posX + Embalses(nroEmb).TamX, _
                    Embalses(nroEmb).posY + 200 - Embalses(nroEmb).TamX * Embalses(nroEmb).porcFD2 / 100, _
                    60, 30)
        Embalses(nroEmb).objPorcFD2.TextFrame.Characters.Text = CStr(Round(Embalses(nroEmb).porcFD2, 1)) & " %"
        If nep(Embalses(nroEmb).nombre, FechaReal + 2) * 100 > Embalses(nroEmb).porcFD2 Then
            Embalses(nroEmb).objPorcFD2.TextFrame.Characters.Font.Color = vbRed
        Else
            Embalses(nroEmb).objPorcFD2.TextFrame.Characters.Font.Color = vbBlack
        End If
        Embalses(nroEmb).objPorcFD2.TextFrame.Characters.Font.Bold = True
        Embalses(nroEmb).objPorcFD2.Fill.Transparency = 0.9
        Embalses(nroEmb).objPorcFD2.Line.Visible = msoFalse
        Embalses(nroEmb).objPorcFD2.TextFrame.AutoSize = True
        
        'Mostrar nivel NEP
        Set Embalses(nroEmb).objNivelNEP = HojaGrafico.Shapes.AddShape(msoShapeRectangle, Embalses(nroEmb).posX, Embalses(nroEmb).posY + (200 - Embalses(nroEmb).TamX * nep(Embalses(nroEmb).nombre, FechaReal)), Embalses(nroEmb).TamX, Embalses(nroEmb).TamX * nep(Embalses(nroEmb).nombre, FechaReal))
        Embalses(nroEmb).objNivelNEP.Fill.Patterned msoPatternOutlinedDiamond
        Embalses(nroEmb).objNivelNEP.Fill.ForeColor.RGB = RGB(50, 50, 50)
        Embalses(nroEmb).objNivelNEP.Fill.Transparency = 0.7
        
        'Mostrar Etiqueta NEP
        Set Embalses(nroEmb).objValorNEP = HojaGrafico.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                    Embalses(nroEmb).posX + Embalses(nroEmb).TamX - 45, _
                    Embalses(nroEmb).posY + 200 - Embalses(nroEmb).TamX, _
                    60, 30)
        Embalses(nroEmb).objValorNEP.TextFrame.Characters.Text = "NEP " & Chr(10) & CStr(Round(Embalses(nroEmb).nep * 100, 1)) & " %"
        Embalses(nroEmb).objValorNEP.TextFrame.Characters.Font.Color = vbBlack
        Embalses(nroEmb).objValorNEP.TextFrame.Characters.Font.Bold = True
        Embalses(nroEmb).objValorNEP.TextFrame.Characters.Font.Size = 11
        Embalses(nroEmb).objValorNEP.Fill.Transparency = 0.9
        Embalses(nroEmb).objValorNEP.Line.Visible = msoFalse
        Embalses(nroEmb).objValorNEP.TextFrame.AutoSize = True


 
        'Mostrar aportes de embalses
        Dim porc As Single
        Dim m3s As Single
        porc = AportesEmbalse(Embalses(nroEmb).nombre)
        If porc = 0 Then porc = 100
        m3s = AportesEmbalse(Embalses(nroEmb).nombre, True)
        Set Embalses(nroEmb).objRiosPorc = HojaGrafico.Shapes.AddShape(msoShapeRectangle, _
            Embalses(nroEmb).posX, Embalses(nroEmb).posY + 170 - Embalses(nroEmb).TamX, _
            Embalses(nroEmb).TamX * porc / 200, 20)
        Embalses(nroEmb).objRiosPorc.TextFrame.Characters.Text = CStr(Round(m3s, 0)) & " m3/s"
        Embalses(nroEmb).objRiosPorc.TextFrame.Characters.Font.Color = vbBlack
        Embalses(nroEmb).objRiosPorc.TextFrame.Characters.Font.Bold = True
        Embalses(nroEmb).objRiosPorc.TextFrame.Characters.Font.Size = 10
        Embalses(nroEmb).objRiosPorc.Fill.Transparency = 0.65

        'Mostrar generacion de central
        Dim GenCentral As Single
        Dim PorcGenCen As Single
        GenCentral = InfoBal(nmCen, "Generacion GWh/dia", "CENTRAL", "GWh/dia", "hoy")
        If Embalses(nroEmb).capCenGWhd > 0 Then
            PorcGenCen = GenCentral / Embalses(nroEmb).capCenGWhd
            Set Embalses(nroEmb).objGenPorc = HojaGrafico.Shapes.AddShape(msoShapeRectangle, _
                Embalses(nroEmb).posX, Embalses(nroEmb).posY + 210, _
                Embalses(nroEmb).TamX * PorcGenCen, 20)
            Embalses(nroEmb).objGenPorc.TextFrame.Characters.Text = CStr(Round(GenCentral, 1)) & " GWh"
            Embalses(nroEmb).objGenPorc.TextFrame.Characters.Font.Color = vbBlack
            Embalses(nroEmb).objGenPorc.TextFrame.Characters.Font.Bold = True
            Embalses(nroEmb).objGenPorc.TextFrame.Characters.Font.Size = 10
            Embalses(nroEmb).objGenPorc.Fill.ForeColor.RGB = RGB(50, 50, 50)
            Embalses(nroEmb).objGenPorc.Fill.Transparency = 0.65
        End If


        filaEmb = filaEmb + 1
        nmEmbalse = UCase(Trim(ThisWorkbook.Worksheets("Embalses").Cells(filaEmb, ColEmbEmbalseIDO).Value))
    Loop
    
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " DibujarEmbalses"
End Sub


'Lee y devuelve información del embalse obtenida en la hoja balance
'dia: hoy, sig
'IoF: I-inicial, F-Final
'Uso x = InfoBal("guavio","Energia disponible","embalse","GWh","hoy")
Public Function InfoBal(nm As String, info As String, Tipo As String, unidad As String, dia As String) As Single

    Dim fila As Integer
    Dim colDia As Integer
    Dim blnHallado As Boolean
    Dim strNm As String
    Dim strInfo As String
    Dim strDia As String
    Dim strTipo As String
    Dim strUnidad As String
    
    nm = UCase(Trim(nm))
    info = UCase(Trim(info))
    dia = UCase(Trim(dia))
    Tipo = UCase(Trim(Tipo))
    unidad = UCase(Trim(unidad))
    
    If dia = "HOY" Then colDia = 4
    If dia = "SIG" Then colDia = 5
    
    blnHallado = False
    fila = 1
    
    strInfo = UCase(Trim(ThisWorkbook.Worksheets("Balances").Cells(fila, 1).Value))
    strUnidad = UCase(Trim(ThisWorkbook.Worksheets("Balances").Cells(fila, 2).Value))
    strTipo = UCase(Trim(ThisWorkbook.Worksheets("Balances").Cells(fila, 6).Value))
    strNm = UCase(Trim(ThisWorkbook.Worksheets("Balances").Cells(fila, 7).Value))
    
    Do While blnHallado = False And fila < 1000
        DoEvents
        If strNm = nm And strInfo = info And Tipo = strTipo And strUnidad = unidad Then
            blnHallado = True
            InfoBal = ThisWorkbook.Worksheets("Balances").Cells(fila, colDia).Value
        End If
    
        fila = fila + 1
        strInfo = UCase(Trim(ThisWorkbook.Worksheets("Balances").Cells(fila, 1).Value))
        strUnidad = UCase(Trim(ThisWorkbook.Worksheets("Balances").Cells(fila, 2).Value))
        strTipo = UCase(Trim(ThisWorkbook.Worksheets("Balances").Cells(fila, 6).Value))
        strNm = UCase(Trim(ThisWorkbook.Worksheets("Balances").Cells(fila, 7).Value))
    
    Loop
    
    
    

End Function







