Option Explicit


Public Const colEquivAgente = 1
Public Const colEquivTipo = 2
Public Const colEquivCentralDDEC = 3
Public Const colEquivCentralOFEI = 4
Public Const colEquivCentralIDO = 5
Public Const colEquivCentralPreideal = 6
Public Const colEquivInformeGenProg = 7
Public Const colEquivInformeGenRealEmp = 8
Public Const colEquivEmbBalance = 9

Public Const filaEquivInicio = 2

Public Type typeEquiv
    Agente As String
    Tipo As String
    CentralDDEC As String
    centralOFEI As String
    CentralIDO As String
    centralPreIdeal As String
    informeGenProg As String
    informeGenRealEmp As String
    EmbBalance As String
End Type

Public Equivalencias(MAXCENTRALES) As typeEquiv
Public nroEquiv As Integer

Public Sub LeerEquivalencias()
    Dim FilaEquiv As Integer
    Dim Agente As String
    Dim central As String
    Dim i As Integer
    
    i = 0
    FilaEquiv = filaEquivInicio
    Agente = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivAgente).Value))
    Do While Agente <> ""
        DoEvents
        i = i + 1
        Equivalencias(i).Agente = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivAgente).Value))
        Equivalencias(i).Tipo = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivTipo).Value))
        Equivalencias(i).CentralDDEC = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivCentralDDEC).Value))
        Equivalencias(i).centralOFEI = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivCentralOFEI).Value))
        Equivalencias(i).CentralIDO = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivCentralIDO).Value))
        Equivalencias(i).centralPreIdeal = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivCentralPreideal).Value))
        Equivalencias(i).informeGenProg = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivInformeGenProg).Value))
        Equivalencias(i).informeGenRealEmp = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivInformeGenRealEmp).Value))
        Equivalencias(i).EmbBalance = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivEmbBalance).Value))
        FilaEquiv = FilaEquiv + 1
        Agente = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(FilaEquiv, colEquivAgente).Value))
        
    Loop
    nroEquiv = i
End Sub

Public Function CenOFEIaDDEC(nmCenOFEI As String) As String
    Dim fila As Integer
    Dim blnHallada As Boolean
    Dim strCenOFEI As String
    
    
    nmCenOFEI = UCase(Trim(nmCenOFEI))
    fila = 2
    CenOFEIaDDEC = ""
    strCenOFEI = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(fila, colEquivCentralOFEI).Value))
    Do While blnHallada = False And fila < MAXCENTRALES
        DoEvents
        If strCenOFEI = nmCenOFEI Then
            CenOFEIaDDEC = ThisWorkbook.Worksheets("Equivalencias").Cells(fila, colEquivCentralDDEC).Value
        End If
        fila = fila + 1
        strCenOFEI = UCase(Trim(ThisWorkbook.Worksheets("Equivalencias").Cells(fila, colEquivCentralOFEI).Value))
    Loop


End Function
