Sub ActualizarStatusProject_V2()
    ' ==========================================
    ' CONFIGURACIÓN
    ' ==========================================
    Dim strRutaExcel As String
    Dim nombreHoja As String
    Dim columnaDatos As String
    Dim i As Integer
    Dim filaInicial As Integer
    
    ' Ruta del archivo (He mantenido la corrección de 'Docuemnts' a 'Documents')
    strRutaExcel = "C:\users\x373176\Documents\Emilio\2026_ONB_Follow Up.xlsx"
    
    ' Nombre de la pestaña
    nombreHoja = "1_GENERAL_STATUS_GLOBAL"
    
    ' CAMBIO REALIZADO: Ahora apunta a la columna C
    columnaDatos = "C"
    
    ' CAMBIO REALIZADO: Fila 3 (porque la 2 suelen ser encabezados en tu imagen)
    ' Si ves que se salta el primero, cámbialo a 2.
    filaInicial = 3
    
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
    
    ' Verificación de seguridad de archivo
    If Dir(strRutaExcel) = "" Then
        MsgBox "No encuentro el archivo. Verifica que la ruta y el nombre sean exactos:" & vbCrLf & strRutaExcel, vbCritical
        Exit Sub
    End If

    On Error Resume Next
    
    ' Abrir Excel (Invisible)
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(strRutaExcel, ReadOnly:=True)
    Set xlSheet = xlBook.Sheets(nombreHoja)
    
    If Err.Number <> 0 Then
        MsgBox "Error al abrir Excel. ¿La pestaña '" & nombreHoja & "' existe?", vbCritical
        xlBook.Close False
        xlApp.Quit
        Exit Sub
    End If
    On Error GoTo 0

    ' ==========================================
    ' BUCLE DE ACTUALIZACIÓN
    ' ==========================================
    
    ' Define la diapositiva actual (la que tengas en pantalla al ejecutar)
    Set sld = ActiveWindow.View.Slide
    
    ' Iteramos 12 veces (Epic 1 al 12)
    For i = 1 To 12
        ' Nombre de la forma en PPT: Epic_Desc_1, Epic_Desc_2...
        nombreForma = "Epic_Desc_" & i
        
        ' Dato de Excel: Columna C, Fila variable
        ' Fila = 3 + (0) = 3 para el primero
        ' Fila = 3 + (1) = 4 para el segundo...
        datoExcel = xlSheet.Range(columnaDatos & (filaInicial + (i - 1))).Value
        
        ' Buscar y reemplazar
        On Error Resume Next
        Set shp = sld.Shapes(nombreForma)
        
        If Not shp Is Nothing Then
            ' Inyectar texto manteniendo formato
            shp.TextFrame.TextRange.Text = datoExcel
        Else
            ' Si no encuentra la forma, lo avisa en la ventana "Inmediato" (Ctrl+G)
            Debug.Print "No encuentro la forma: " & nombreForma
        End If
        On Error GoTo 0
        
        Set shp = Nothing
    Next i

    ' ==========================================
    ' CERRAR Y LIMPIAR
    ' ==========================================
    xlBook.Close SaveChanges:=False
    xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    MsgBox "Datos de la Columna C actualizados correctamente.", vbInformation

End Sub
