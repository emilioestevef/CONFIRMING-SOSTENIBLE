Sub ActualizarNombresEpicas()
    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim rutaArchivo As String
    Dim f As Integer
    Dim nombreEpica As String
    
    ' --- RUTA ACTUALIZADA ---
    rutaArchivo = "C:\users\x373176\Docuemnts\Emilio\2026_ONB_Follow Up.xlsx"
    ' ------------------------

    On Error Resume Next
    ' Intentamos conectar con Excel
    Set excelApp = GetObject(, "Excel.Application")
    If excelApp Is Nothing Then
        Set excelApp = CreateObject("Excel.Application")
        excelApp.Visible = True
    End If
    On Error GoTo 0

    ' Abrir el archivo de Emilio
    On Error Resume Next
    Set wb = excelApp.Workbooks("2026_ONB_Follow Up.xlsx") ' Si ya está abierto
    If wb Is Nothing Then
        Set wb = excelApp.Workbooks.Open(rutaArchivo) ' Si hay que abrirlo
    End If
    
    If wb Is Nothing Then
        MsgBox "No he podido encontrar el archivo. Revisa si la carpeta se escribe 'Documents' o 'Docuemnts'.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Apuntamos a la hoja correcta
    Set ws = wb.Sheets("1_GENERAL_STATUS_GLOBAL")

    ' Bucle para las 12 cajas de tu PPT
    For f = 1 To 12
        ' Tomamos el valor de la columna C (3), empezando en fila 3 (f+2)
        nombreEpica = ws.Cells(f + 2, 3).Value 
        
        On Error Resume Next
        ' Esto escribe el nombre en tus cajas azules llamadas Epic_Desc_1, etc.
        With ActivePresentation.Slides(1).Shapes("Epic_Desc_" & f).TextFrame.TextRange
            .Text = nombreEpica
            .Font.Name = "Santander Text" ' Mantenemos tu fuente corporativa
            .Font.Size = 10
            .Font.Color.RGB = RGB(255, 255, 255) ' Texto en blanco
        End With
        On Error GoTo 0
    Next f

    MsgBox "¡Nombres de Épicas actualizados con éxito!", vbInformation, "Onboarding Status"
End Sub
