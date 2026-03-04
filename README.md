Sub ActualizarStatusProject()
    ' ==========================================
    ' CONFIGURACIÓN (REVISA ESTO SI FALLA)
    ' ==========================================
    Dim strRutaExcel As String
    Dim nombreHoja As String
    Dim columnaDatos As String
    Dim i As Integer
    Dim filaInicial As Integer
    
    ' CORRECCIÓN: He corregido "Docuemnts" a "Documents" en la ruta.
    ' Si tu carpeta realmente se llama "Docuemnts" (con error), cámbialo abajo.
    strRutaExcel = "C:\users\x373176\Documents\Emilio\2026_ONB_Follow Up.xlsx"
    
    ' Nombre exacto de la pestaña del Excel (según tu captura de pantalla)
    nombreHoja = "1_GENERAL_STATUS_GLOBAL" 
    
    ' Columna donde está el texto (B es la columna 2)
    columnaDatos = "B" 
    
    ' La fila donde empieza el primer dato (Epic #1)
    filaInicial = 2 
    
    ' ==========================================
    ' VARIABLES DEL SISTEMA
    ' ==========================================
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim sld As Slide
    Dim shp As Shape
    Dim datoExcel As String
    Dim nombreForma As String
    
    ' Comprobamos si el archivo existe antes de empezar
    If Dir(strRutaExcel) = "" Then
        MsgBox "No encuentro el archivo Excel en la ruta: " & vbCrLf & strRutaExcel, vbCritical
        Exit Sub
    End If

    On Error Resume Next
    
    ' Abrir Excel en modo invisible (silencioso)
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(strRutaExcel, ReadOnly:=True) ' Solo lectura para no bloquear
    Set xlSheet = xlBook.Sheets(nombreHoja)
    
    If Err.Number <> 0 Then
        MsgBox "Error al abrir el Excel. Verifica el nombre de la hoja.", vbCritical
        xlBook.Close False
        xlApp.Quit
        Exit Sub
    End If
    On Error GoTo 0

    ' ==========================================
    ' EL BUCLE MAGICO (Actualiza las slides)
    ' ==========================================
    
    ' Definimos en qué diapositiva estamos trabajando (Diapositiva Actual)
    ' Si siempre es la diapositiva 3, cambia ActiveWindow.View.Slide por ActivePresentation.Slides(3)
    Set sld = ActiveWindow.View.Slide
    
    ' Iteramos 12 veces (para los 12 Epics)
    For i = 1 To 12
        ' Construimos el nombre de la forma: Epic_Desc_1, Epic_Desc_2...
        nombreForma = "Epic_Desc_" & i
        
        ' Obtenemos el dato de Excel
        ' Fila = filaInicial + (i - 1). 
        ' Ej: i=1 -> Fila 2. i=2 -> Fila 3.
        datoExcel = xlSheet.Range(columnaDatos & (filaInicial + (i - 1))).Value
        
        ' Buscamos la forma en la slide y actualizamos
        On Error Resume Next
        Set shp = sld.Shapes(nombreForma)
        
        If Not shp Is Nothing Then
            ' Actualizamos SOLO el texto, manteniendo el formato (color, fuente, tamaño)
            shp.TextFrame.TextRange.Text = datoExcel
        Else
            Debug.Print "No encontré la forma: " & nombreForma
        End If
        On Error GoTo 0
        
        ' Reseteamos la variable de la forma para la siguiente vuelta
        Set shp = Nothing
    Next i

    ' ==========================================
    ' LIMPIEZA (Cerrar Excel)
    ' ==========================================
    xlBook.Close SaveChanges:=False
    xlApp.Quit
    
    ' Liberar memoria
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    MsgBox "¡Actualización completada con éxito!", vbInformation

End Sub
