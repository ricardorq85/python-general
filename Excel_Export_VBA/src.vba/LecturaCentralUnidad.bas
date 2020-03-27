Public Const colPlaUnidad = 1
Public Const colPlaCentral = 2


Public Type typeCenDeUnidad
    unidad As String
    central As String
End Type

Public CentralDeUnidad(MAXUNIDADES) As typeCenDeUnidad



Public Function LeerCentralUnidad() As Integer
    Dim i As Integer
    Dim strUnidad As String
    Dim strCentral As String
    Dim fila As String
    
    i = 1
    fila = 2
    strUnidad = ThisWorkbook.Worksheets("PlantaUnidad").Cells(fila, colPlaUnidad).Value
    strCentral = ThisWorkbook.Worksheets("PlantaUnidad").Cells(fila, colPlaCentral).Value
    Do While strUnidad <> ""
        DoEvents
        CentralDeUnidad(i).unidad = UCase(Trim(strUnidad))
        CentralDeUnidad(i).central = UCase(Trim(strCentral))
    
        fila = fila + 1
        i = i + 1
        strUnidad = ThisWorkbook.Worksheets("PlantaUnidad").Cells(fila, colPlaUnidad).Value
        strCentral = ThisWorkbook.Worksheets("PlantaUnidad").Cells(fila, colPlaCentral).Value
    Loop
    
    LeerCentralUnidad = i - 1
    
    OrdenarCentralDeUnidadPorUnidad CentralDeUnidad
    
    
End Function

Sub OrdenarCentralDeUnidadPorUnidad(CenUni() As typeCenDeUnidad)
    
    Dim nroUnidad As Integer
    Dim nmUnidad As String

    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    nroUnidad = 1
    nmUnidad = CenUni(nroUnidad).unidad
    Do While nmUnidad <> ""
        DoEvents
        nroUnidad = nroUnidad + 1
        nmUnidad = CenUni(nroUnidad).unidad
    Loop
    nroUnidad = nroUnidad - 1
    
    'Ordenar unidades por nombre
    For i = 1 To nroUnidad - 1
        For j = i + 1 To nroUnidad
            If UCase(CenUni(i).unidad) > UCase(CenUni(j).unidad) Then
            
                CenUni(0).unidad = CenUni(i).unidad
                CenUni(0).central = CenUni(i).central
                
                CenUni(i).unidad = CenUni(j).unidad
                CenUni(i).central = CenUni(j).central

                CenUni(j).unidad = CenUni(0).unidad
                CenUni(j).central = CenUni(0).central
            End If
        Next j
    Next i
    
End Sub


Public Function nmCentralUnidad(nmUnidad As String, CenUni() As typeCenDeUnidad, nroUnidades As Integer) As String

    Dim unidadInf As String
    Dim unidadSup As String
    Dim unidadMed As String
    Dim blnHallado As String
    Dim intSup As Integer
    Dim intInf As Integer
    Dim intMed As Integer
    Dim PosUnidadEnCenUni As Integer
    
    nmCentral = UCase(Trim(nmUnidad))
    
    PosUnidadEnCenUni = -1
    
    intSup = nroUnidades
    intInf = 1
    intMed = Int((intSup + intInf) / 2)
    unidadInf = UCase(Trim(CenUni(intInf).unidad))
    unidadSup = UCase(Trim(CenUni(intSup).unidad))
    unidadMed = UCase(Trim(CenUni(intMed).unidad))
    
    If unidadInf = nmUnidad Then
        PosUnidadEnCenUni = intInf
        Exit Function
    End If
    
    If unidadSup = nmUnidad Then
        PosUnidadEnCenUni = intSup
        Exit Function
    End If
    
    If unidadMed = nmUnidad Then
        PosUnidadEnCenUni = intSup
        Exit Function
    End If
    
    blnHallado = False
    
    Do While unidadMed <> nmUnidad And blnHallado = False And (intSup - intInf > 1)
        
        DoEvents
        
        If nmUnidad > unidadMed Then
            intInf = intMed
            unidadInf = UCase(Trim(CenUni(intInf).unidad))
        End If
        
        If nmUnidad < unidadMed Then
            intSup = intMed
            unidadSup = UCase(Trim(CenUni(intSup).unidad))
        End If
        
        intMed = Int((intSup + intInf) / 2)
        unidadMed = UCase(Trim(CenUni(intMed).unidad))
        
        If nmUnidad = unidadMed Then
            PosUnidadEnCenUni = intMed
            blnHallado = True
        End If
    Loop
    
    If PosUnidadEnCenUni <> -1 Then nmCentralUnidad = CenUni(PosUnidadEnCenUni).central

End Function
    


