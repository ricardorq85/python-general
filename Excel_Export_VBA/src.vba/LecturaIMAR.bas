Option Explicit

Public Type typeIMAR
    MWh(24) As Single
End Type

Public iCostoMarginal As typeIMAR

Sub LeerIMAR(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoIMAR(fecha)
    Dim i As Integer
    Dim Hora As Integer
    Dim Suma As Single
    Dim Mayor As Single
    i = 0
    
    ThisWorkbook.Worksheets("PreIdeal").Cells(2, 1).Value = "Costo Marginal"
    
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
                    iCostoMarginal.MWh(Hora) = LArray(Hora)
                    If iCostoMarginal.MWh(Hora) > Mayor Then Mayor = iCostoMarginal.MWh(Hora)
                    ThisWorkbook.Worksheets("PreIdeal").Cells(2, Hora + 1).Value = iCostoMarginal.MWh(Hora)
                    Suma = Suma + iCostoMarginal.MWh(Hora)
                Next Hora
                ThisWorkbook.Worksheets("PreIdeal").Cells(2, Hora + 1).Value = Suma / 24
                ThisWorkbook.Worksheets("PreIdeal").Cells(2, Hora + 2).Value = Mayor
                i = i + 1
            End If
        Loop
    Close #1
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerIMAR"
End Sub


Public Function ArchivoIMAR(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamIMAR, ColParamPrefijo).Value
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoIMAR = raiz & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & "_NAL.txt"
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamIMAR, ColParamRaiz).Value
        ArchivoIMAR = raiz & año & "\" & mes & "\" & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & "_NAL.txt"
    End If
End Function






