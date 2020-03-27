Option Explicit

Public Type typeDAGC
    central As String
    MWh(24) As Single
End Type

Public DespachoAGC(MAXCENTRALES) As typeDAGC

Sub LeerDAGC(fecha As Date, Optional posX As Integer = 0, Optional posY As Integer = 0)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoDAGC(fecha)
    Dim nroCenAGC As Integer
    Dim j As Integer
    Dim Hora As Integer
    Dim nmCen As String
    nroCenAGC = 0
    
    ThisWorkbook.Worksheets("Servicio AGC").UsedRange.Delete
    ThisWorkbook.Worksheets("Servicio AGC").Cells(1, 1).Value = fecha
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 24 Then
                'Debug.Print textline
                nroCenAGC = nroCenAGC + 1
                DespachoAGC(nroCenAGC).central = Trim(LArray(0))
                For Hora = 1 To 24
                    DespachoAGC(nroCenAGC).MWh(Hora) = LArray(Hora)
                Next Hora
                
            End If
        Loop
    Close #1
    
    Dim TotalCentral As Double
    
    ThisWorkbook.Worksheets("Servicio AGC").Cells(1 + posX, 1 + posY).Value = fecha
    ThisWorkbook.Worksheets("Servicio AGC").Cells(2 + posX, 1 + posY).Value = "Hora"
    ThisWorkbook.Worksheets("Servicio AGC").Cells(1 + posX, 1 + posY).Value = fecha
    ThisWorkbook.Worksheets("Servicio AGC").Cells(1 + posX, 2 + posY).Value = "Asignación de AGC por planta(MWh)"
    ThisWorkbook.Worksheets("Servicio AGC").Activate
    ThisWorkbook.Worksheets("Servicio AGC").Range(Cells(1 + posX, 1 + posY), Cells(1 + posX, 1 + posY + nroCenAGC)).Interior.Color = RGB(170, 170, 170)
    ThisWorkbook.Worksheets("Servicio AGC").Range(Cells(2 + posX, 1 + posY), Cells(2 + posX, 1 + posY + nroCenAGC)).Interior.Color = RGB(200, 200, 200)
    ThisWorkbook.Worksheets("Servicio AGC").Range(Cells(3 + posX, 1 + posY), Cells(1 + posX + 26, 1 + posY + nroCenAGC)).Interior.Color = RGB(230, 230, 230)
    ThisWorkbook.Worksheets("Servicio AGC").Range(Cells(1 + posX + 26, 1 + posY), Cells(1 + posX + 26, 1 + posY + nroCenAGC)).Interior.Color = RGB(200, 200, 200)
    
    For Hora = 1 To 24
        ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2 + posX, 1 + posY).Value = Hora
        ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2 + posX, 1 + posY).Interior.Color = RGB(200, 200, 200)
    Next Hora
    ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2 + posX, 1 + posY).Value = "TOTAL"




    
    For j = 1 To nroCenAGC
        nmCen = EliminarComillas(DespachoAGC(j).central)
        ThisWorkbook.Worksheets("Servicio AGC").Cells(2, j + 1).Value = nmCen
        ThisWorkbook.Worksheets("Informe").Cells(130 + j, 2).Value = nmCen
        TotalCentral = 0
        For Hora = 1 To 24
            ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2, j + 1).Value = DespachoAGC(j).MWh(Hora)
            TotalCentral = TotalCentral + DespachoAGC(j).MWh(Hora)
        Next Hora
        ThisWorkbook.Worksheets("Servicio AGC").Cells(Hora + 2, j + 1).Value = TotalCentral
        ThisWorkbook.Worksheets("Informe").Cells(130 + j, 3).Value = TotalCentral
    Next j
    
    For j = nroCenAGC + 1 To 12
        ThisWorkbook.Worksheets("Informe").Cells(130 + j, 2).Value = ""
        ThisWorkbook.Worksheets("Informe").Cells(130 + j, 3).Value = ""
    Next j
    
    LeerOfeiAgcPlanta fecha, 0, nroCenAGC + 2
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerDAGC"
       
End Sub


Public Function ArchivoDAGC(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamDAGC, ColParamPrefijo).Value
    año = Year(fecha)
    
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoDAGC = raiz & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamDAGC, ColParamRaiz).Value
        ArchivoDAGC = raiz & año & "\" & mes & "\" & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & ".txt"
    End If
End Function








