Public Sub RevisarPreciosMarginales(fecha As Date, nmHoja As String)
    Dim fila As Integer
    Dim nmCentral As String
    Dim Hora As Integer
    Dim MaxQ As Single
    Dim MinQ As Single
    Dim MaxP As Single
    Dim MinP As Single
    Dim D As Single 'Disponibilidad
    Dim Q As Single
    Dim P As Single
    Dim MO As Single
    Dim AGC As Single
    Dim MX As Single
    Dim MW As Single
    Dim Total As Single
    Dim sumaQ As Single
    Dim nQ As Single
    Dim promQ As Single
    Dim codHashCen As Integer
    
    LeerOfeiDispPlanta fecha, False
    If nmHoja = "DDEC" Then CargarAGC_DispCen fecha
    If nmHoja = "DDEC" Then CargarSEGDES_DispCen fecha
    
    fila = 3
    nmCentral = UCase(Trim(ThisWorkbook.Worksheets(nmHoja).Cells(fila, 1).Value))
    Do While nmCentral <> "" And fila < MAXCENTRALES
        DoEvents
        
        If nmCentral <> "" And CentralEsMayor(nmCentral) Then
            'Revisar cantidades
            MaxQ = -1
            MinQ = 10000
            
            For Hora = 1 To 24
                Q = ThisWorkbook.Worksheets(nmHoja).Cells(fila, Hora + 1).Value
                Total = Total + Q
                If Q > MaxQ Then MaxQ = Q
                If (Q < MinQ) Then MinQ = Q
            Next Hora

            If MaxQ <> MinQ Then
                
                codHashCen = CodigoHash(nmCentral)
            
            
                MaxP = -1
                MinP = 10000000
                For Hora = 1 To 24
                    D = DispCen(codHashCen).MWh(Hora)
                    P = ThisWorkbook.Worksheets(nmHoja).Cells(2, Hora + 1).Value
                    Q = ThisWorkbook.Worksheets(nmHoja).Cells(fila, Hora + 1).Value
                    MO = DispCen(codHashCen).MO(Hora)
                    If nmHoja = "DDEC" Then AGC = DispCen(codHashCen).AGC(Hora) Else AGC = 0
                    If nmHoja = "DDEC" Then MX = DispCen(codHashCen).MX(Hora) Else MX = 99999
                    If nmHoja = "DDEC" Then MW = DispCen(codHashCen).MW(Hora) Else MW = 99999
                    If MW = 0 Then MW = 99999
                    If MX = 0 Then MX = 99999
                    If (Q + AGC) < MIN(D, MX) And Q > 3 And Q <> MO And Q <> MW Then
                        If P > MaxP Then MaxP = P
                        If P < MinP Then MinP = P
                    End If
                Next Hora
            
                'Calcular promedio de Q
                sumaQ = 0
                nQ = 0
                promQ = 0
                For Hora = 1 To 24
                    D = DispCen(codHashCen).MWh(Hora)
                    P = ThisWorkbook.Worksheets(nmHoja).Cells(2, Hora + 1).Value
                    Q = ThisWorkbook.Worksheets(nmHoja).Cells(fila, Hora + 1).Value
                    MO = DispCen(codHashCen).MO(Hora)
                    If nmHoja = "DDEC" Then AGC = DispCen(codHashCen).AGC(Hora) Else AGC = 0
                    If nmHoja = "DDEC" Then MX = DispCen(codHashCen).MX(Hora) Else MX = 99999
                    If nmHoja = "DDEC" Then MW = DispCen(codHashCen).MW(Hora) Else MW = 99999
                    If MW = 0 Then MW = 99999
                    If MX = 0 Then MX = 99999
                    If (Q + AGC) < MIN(D, MX) And Q > 3 And Q <> MO And Q <> MW Then
                        If P = MinP Then
                            sumaQ = sumaQ + Q
                            nQ = nQ + 1
                        End If
                    End If
                Next Hora
                If nQ > 0 Then promQ = sumaQ / nQ

                'Marcar centrales
                For Hora = 1 To 24
                    D = DispCen(codHashCen).MWh(Hora)
                    MO = DispCen(codHashCen).MO(Hora)
                    If nmHoja = "DDEC" Then AGC = DispCen(codHashCen).AGC(Hora) Else AGC = 0
                    If nmHoja = "DDEC" Then MX = DispCen(codHashCen).MX(Hora) Else MX = 99999
                    If nmHoja = "DDEC" Then MW = DispCen(codHashCen).MW(Hora) Else MW = 99999
                    If MW = 0 Then MW = 99999
                    If MX = 0 Then MX = 99999
                    P = ThisWorkbook.Worksheets(nmHoja).Cells(2, Hora + 1).Value
                    Q = ThisWorkbook.Worksheets(nmHoja).Cells(fila, Hora + 1).Value
                    If (Q + AGC) < MIN(D, MX) And Q > 3 Then
                        If P = MinP And (Q <> promQ Or nQ = 1) And Q <> MO And Q <> MW Then
                            ThisWorkbook.Worksheets(nmHoja).Cells(fila, Hora + 1).Interior.Color = vbCyan
                            ThisWorkbook.Worksheets(nmHoja).Cells(3, Hora + 1).Value = nmCentral
                        End If
                    End If
                Next Hora

            End If
        
        End If
        
        
        fila = fila + 1
        nmCentral = UCase(Trim(ThisWorkbook.Worksheets(nmHoja).Cells(fila, 1).Value))
    Loop
    
    If nmHoja = "DDEC" Then RevisarPlantasMarginalesDDEC
    
End Sub

Sub CargarAGC_DispCen(fecha As Date)
    Dim archivo As String
    Dim textline As String
    Dim LArray() As String
    archivo = ArchivoDAGC(fecha)
    Dim Hora As Integer
    Dim nmCen As String
    Dim codHashCen As Integer
    'nroCenAGC = 0
    
    On Error GoTo ManejadorError
    Open archivo For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            LArray = Split(textline, ",")
            If UBound(LArray) = 24 Then
                
                nmCen = EliminarComillas(UCase(Trim(LArray(0))))
                codHashCen = CodigoHash(nmCen)
                For Hora = 1 To 24
                    DispCen(codHashCen).AGC(Hora) = LArray(Hora)
                Next Hora
                'Debug.Print codHashCen; "  "; textline
            End If
        Loop
    Close #1
    Exit Sub
ManejadorError:
    LogOfertaEPM Err.Description & " " & archivo & " CargarAGC_DispCen"
End Sub

Public Sub RevisarPlantasMarginalesDDEC()
    Dim FilaPrecio As Integer
    Dim FilaPlanta As Integer
    Dim Hora As Integer
    Dim PlantaDDEC As String
    Dim PlantaPrId As String
    Dim PrecioMarDDEC As Single
    FilaPrecio = 2
    FilaPlanta = 3
    
    For Hora = 1 To 24
        PlantaDDEC = UCase(Trim(ThisWorkbook.Worksheets("DDEC").Cells(FilaPlanta, Hora + 1).Value))
        If PlantaDDEC = "" Then  'No se pudo identificar la planta en la revision del DDEC. Se intentara traer el precio del preideal.
            PrecioMarDDEC = Trim(ThisWorkbook.Worksheets("DDEC").Cells(FilaPrecio, Hora + 1).Value)
            PlantaPrId = UCase(Trim(HallarPlantaConPrecio(PrecioMarDDEC)))
            If PlantaPrId <> "" Then 'Se encontro el precio de la planta en el preideal
                ThisWorkbook.Worksheets("DDEC").Cells(FilaPlanta, Hora + 1).Value = PlantaPrId
                ThisWorkbook.Worksheets("DDEC").Cells(FilaPlanta, Hora + 1).Interior.Color = RGB(150, 150, 255) 'Azul opaco
            Else 'No se encontrol el precio de la planta en el preideal
                ThisWorkbook.Worksheets("DDEC").Cells(FilaPlanta, Hora + 1).Interior.Color = RGB(255, 100, 100) 'Rojo opaco
            End If
        Else 'El analisis del DDEC identifico la planta
            PrecioMarDDEC = Trim(ThisWorkbook.Worksheets("DDEC").Cells(FilaPrecio, Hora + 1).Value)
            PlantaPrId = UCase(Trim(HallarPlantaConPrecio(PrecioMarDDEC)))
            If PlantaPrId <> PlantaDDEC Then 'La planta identificada no coincide con el preideal
                If PlantaPrId <> "" Then 'La planta es diferente
                    ThisWorkbook.Worksheets("DDEC").Cells(FilaPlanta, Hora + 1).Value = PlantaPrId
                    ThisWorkbook.Worksheets("DDEC").Cells(FilaPlanta, Hora + 1).Interior.Color = vbGreen
                Else 'No en encontro ninguna planta con ese precio
                    ThisWorkbook.Worksheets("DDEC").Cells(FilaPlanta, Hora + 1).Interior.Color = vbMagenta
                End If
            End If
        End If
    Next Hora
End Sub
'Busca en el Preideal la planta que tenga el precio marginal dado.
Public Function HallarPlantaConPrecio(prmPrecio As Single) As String
    Dim blnHallado As Boolean
    Dim FilaPrecio As Integer
    Dim FilaPlanta As Integer
    Dim Hora As Integer
    Dim Planta As String
    Dim PrecioMarPrIdeal As Single
    FilaPrecio = 2
    FilaPlanta = 3
    HallarPlantaConPrecio = ""
    Hora = 1
    Do While blnHallado = False And Hora <= 24
        PrecioMarPrIdeal = Trim(ThisWorkbook.Worksheets("PreIdeal").Cells(FilaPrecio, Hora + 1).Value)
        If PrecioMarPrIdeal = prmPrecio Then
            blnHallado = True
            Planta = Trim(ThisWorkbook.Worksheets("PreIdeal").Cells(FilaPlanta, Hora + 1).Value)
            HallarPlantaConPrecio = Planta
        End If
        Hora = Hora + 1
    Loop
End Function

