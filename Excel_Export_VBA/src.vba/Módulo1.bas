Sub Macro1()
'
' Macro1 Macro
'

'
    Range("E12:E13").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("E12").Select
End Sub


Sub Modificar1XLCerrado()
'Declaramos variables
Dim archivo As Application
Dim nombreArchivo As String
'
'Creamos el objecto Excel
Set archivo = CreateObject("Excel.Application")
'
With archivo
    '
    'Asignamos el nombre del archivo
    nombreArchivo = "C:\carpeta\Libro1.xlsx"
    '
    'Validamos si el archivo ya está abierto
    If IsFileOpen(nombreArchivo) Then
    Else
        '
        With .Workbooks.Open(nombreArchivo)
            'Hacemos las modificaciones en el archivo
            .Worksheets("Hoja1").Range("A1").Value = "Total1"
            .Worksheets("Hoja1").Range("A2").Value = 11
            'Cerramos el archivo guardando cambios
            .Close SaveChanges:=True
        End With
    End If
    '
    'Cerramos la aplicación de Excel
    .Quit
End With
End Sub

Sub ModificarXLCerrados()
'Declaramos variables
Dim archivo As Application
Dim Celda As Object
Dim nombreArchivo As String
'
'Creamos el objecto Excel
Set archivo = CreateObject("Excel.Application")
'
With archivo
    '
    'Recorremos cada celda de la selección para tomar el nombre de cada archivo
    For Each Celda In Selection
        nombreArchivo = Celda.Value
        '
        'Validamos si el archivo ya está abierto
        If IsFileOpen(nombreArchivo) Then
        Else
            '
            With .Workbooks.Open(nombreArchivo)
                'Hacemos las modificaciones en el archivo
                .Worksheets("Hoja1").Range("A1").Value = "Total"
                .Worksheets("Hoja1").Range("A2").Value = 10
                'Cerramos el archivo guardando cambios
                .Close SaveChanges:=True
            End With
        End If
        '
    Next Celda
    '
    'Cerramos la aplicación de Excel
    .Quit
End With
End Sub

' This function checks to see if a file is open or not. If the file is
' already open, it returns True. If the file is not open, it returns
' False. Otherwise, a run-time error occurs because there is
' some other problem accessing the file.
' Código de macro para comprobar si un archivo ya está abierto
' http://support.microsoft.com/kb/291295/es
'
Function IsFileOpen(filename As String)
Dim filenum As Integer, errnum As Integer
'
On Error Resume Next   ' Turn error checking off.
filenum = FreeFile()   ' Get a free file number.
' Attempt to open the file and lock it.
Open filename For Input Lock Read As #filenum
Close filenum          ' Close the file.
errnum = Err           ' Save the error number that occurred.
On Error GoTo 0        ' Turn error checking back on.
' Check to see which error occurred.
Select Case errnum
    ' No error occurred.
    ' File is NOT already open by another user.
Case 0
    IsFileOpen = False
    ' Error number for "Permission Denied."
    ' File is already opened by another user.
Case 70
    IsFileOpen = True
    ' Another error occurred.
Case Else
    Error errnum
End Select
End Function