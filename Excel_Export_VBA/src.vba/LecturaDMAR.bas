Option Explicit

Public Type typeDMAR
    MWh(24) As Single
End Type

Public CostoMarginal As typeDMAR

Sub LeerDMAR(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoDMAR(fecha)
    Dim i As Integer
    Dim Hora As Integer
    Dim Suma As Single
    Dim Mayor As Single
    Dim Promedio As Single
    i = 0
    
    ThisWorkbook.Worksheets("DDEC").Cells(2, 1).Value = "Costo Marginal"
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 24 Then
                'Debug.Print textline
                Suma = 0
                Mayor = -100
                For Hora = 1 To 24
                    CostoMarginal.MWh(Hora) = LArray(Hora)
                    If CostoMarginal.MWh(Hora) > Mayor Then Mayor = CostoMarginal.MWh(Hora)
                    ThisWorkbook.Worksheets("Precios Generaciones").Cells(Hora + 4, 3).Value = CostoMarginal.MWh(Hora)
                    ThisWorkbook.Worksheets("DDEC").Cells(2, Hora + 1).Value = CostoMarginal.MWh(Hora)
                    Suma = Suma + CostoMarginal.MWh(Hora)
                Next Hora
                Promedio = Suma / 24
                ThisWorkbook.Worksheets("Precios Generaciones").Cells(Hora + 4, 3).Value = Promedio
                ThisWorkbook.Worksheets("DDEC").Cells(2, Hora + 1).Value = Promedio
                ThisWorkbook.Worksheets("DDEC").Cells(2, Hora + 2).Value = Mayor
                i = i + 1
            End If
        Loop
    Close #1
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerDMAR"
End Sub


Public Function ArchivoDMAR(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamDMAR, ColParamPrefijo).Value
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoDMAR = raiz & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamDMAR, ColParamRaiz).Value
        ArchivoDMAR = raiz & año & "\" & mes & "\" & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    End If
End Function









Sub Macro2()
'
' Macro2 Macro
'

'
    Range("F3").Select
    Range("$F$3").SparklineGroups.Add Type:=xlSparkColumnStacked100, SourceData _
        :="B3:E3"
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 9592887
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0
    Range("F10").Select
    Range("$F$10").SparklineGroups.Add Type:=xlSparkColumnStacked100, SourceData _
        :="B10:E10"
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 9592887
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0
    Range("F8").Select
    Range("$F$8").SparklineGroups.Add Type:=xlSparkColumnStacked100, SourceData _
        :="B8:E8"
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 9592887
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0
    Range("F7").Select
    Range("$F$7").SparklineGroups.Add Type:=xlSparkColumn, SourceData:="B6:E6"
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 9592887
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0
    Range("F11").Select
    Range("$F$11").SparklineGroups.Add Type:=xlSparkLine, SourceData:="B11:E11"
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 9592887
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0
    Range("G9").Select
End Sub