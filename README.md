Sub ActualizarNombresEpicas()
    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim f As Integer
    Dim nombreEpica As String
    Dim nombreForma As String

    ' 1. Intentar conectar con el Excel que tienes abierto
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    If excelApp Is Nothing Then
        MsgBox "Por favor, abre el archivo Excel primero.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Apuntamos a tu hoja específica (basado en tu captura)
    Set wb = excelApp.ActiveWorkbook
    Set ws = wb.Sheets("1_GENERAL_STATUS_GLOBAL")

    ' 2. Bucle para las 12 Epics que tienes en la slide
    For f = 1 To 12
        ' En Excel: Fila 3 es f=1 (f+2), Columna C es la 3
        nombreEpica = ws.Cells(f + 2, 3).Value 
        
        ' Nombre de la forma que pusiste en el Panel de Selección
        nombreForma = "Epic_Desc_" & f
        
        ' Actualizar el texto
        On Error Resume Next
        ActivePresentation.Slides(1).Shapes(nombreForma).TextFrame.TextRange.Text = nombreEpica
        
        ' Opcional: Ajustar el formato para que el texto sea blanco y centrado
        With ActivePresentation.Slides(1).Shapes(nombreForma).TextFrame.TextRange
            .Font.Color.RGB = RGB(255, 255, 255) ' Blanco
            .Font.Size = 10 ' Ajusta según necesites
        End With
        On Error GoTo 0
    Next f

    MsgBox "¡Nombres de Épicas volcados con éxito!", vbInformation
End Sub
