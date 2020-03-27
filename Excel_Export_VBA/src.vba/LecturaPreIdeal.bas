Option Explicit

Public Type typePreIdeal
    central As String
    Total As Single
    MWh(24) As Single
End Type

Public Preideal(MAXCENTRALES) As typePreIdeal

Sub LeerPreIdeal(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoPreideal(fecha)
    Dim i As Integer
    Dim Hora As Integer
    i = 0
    
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 24 Then
                'Debug.Print textline
                Preideal(i).central = Trim(LArray(0))
                For Hora = 1 To 24
                    Preideal(i).MWh(Hora) = LArray(Hora)
                    Preideal(i).Total = Preideal(i).Total + Preideal(i).MWh(Hora)
                Next Hora
                i = i + 1
            End If
        Loop
    Close #1
    
End Sub


Public Function ArchivoPreideal(fecha As Date) As String
    Dim raiz As String
    Dim prefijo As String
    Dim año As String
    Dim mes As String
    
    prefijo = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamPrid, ColParamPrefijo).Value
    año = Year(fecha)
    mes = NombreMes(nmMes.largo, fecha)
    If blnUsarRutaAlterna Then
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamRutaAlterna, ColParamRaiz).Value
        ArchivoPreideal = raiz & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & "_NAL.txt"
    Else
        raiz = ThisWorkbook.Worksheets("Parametros").Cells(FilaParamPrid, ColParamRaiz).Value
        ArchivoPreideal = raiz & año & "\" & mes & "\" & prefijo & NombreMes(NumeroConCero, fecha) & nroDia(ConCero, fecha) & "_NAL.txt"
    End If
End Function


Public Sub LeerTxtPreIdeal(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    Dim fila As Integer
    Dim Hora As Integer
    Dim Total As Single
    
    
    
    Application.Calculation = xlCalculationManual
    
    ThisWorkbook.Worksheets("PreIdeal").Columns("A:Z").ClearContents
    
    archivo = ArchivoPreideal(fecha)
    ThisWorkbook.Worksheets("PreIdeal").Cells(1, 1).Value = "PreIdeal  " & fecha
    ThisWorkbook.Worksheets("PreIdeal").Cells(3, 1).Value = "Central"
    For Hora = 1 To 24
        ThisWorkbook.Worksheets("PreIdeal").Cells(1, Hora + 1).Value = "Hora " & CStr(Hora)

    Next Hora
    ThisWorkbook.Worksheets("PreIdeal").Cells(1, 26).Value = "Promedio"
    ThisWorkbook.Worksheets("PreIdeal").Cells(1, 27).Value = "Máximo"
    ThisWorkbook.Worksheets("PreIdeal").Cells(3, 26).Value = "Total"

    fila = 4
    
    'Carga Preideal
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 24 Then
                'Debug.Print textline
                Total = 0
                ThisWorkbook.Worksheets("PreIdeal").Cells(fila, 1).Value = EliminarComillas(LArray(0))
                For Hora = 1 To 24
                    ThisWorkbook.Worksheets("PreIdeal").Cells(fila, Hora + 1).Value = LArray(Hora)
                    Total = Total + LArray(Hora)
                Next Hora
                ThisWorkbook.Worksheets("PreIdeal").Cells(fila, Hora + 1).Value = Total
                fila = fila + 1
            End If
        Loop
    Close #1
    
    LeerIMAR fecha
    FormatoSimpleHoja "PreIdeal"
    RevisarPreciosMarginales fecha, "Preideal"
    
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " LeerTxtPreIdeal"
End Sub

Public Function UltimoCaso() As Integer
    Dim caso  As Integer
    Dim strCaso As String
    
    
    caso = 0
    strCaso = UCase(Trim(ThisWorkbook.Worksheets("Sensibilidades").Cells(2, caso + 3).Value))
    Do While strCaso <> ""
        DoEvents
        
        caso = caso + 1
        strCaso = UCase(Trim(ThisWorkbook.Worksheets("Sensibilidades").Cells(2, caso + 3).Value))
    Loop
    UltimoCaso = caso - 1
End Function












