' ═══════════════════════════════════════════════════════════════════════════
' AUTOMATIZACIÓN CRUZ ROJA - CON FILTRO DE COLUMNAS
' Mantiene solo las columnas especificadas y elimina las demás
' ═══════════════════════════════════════════════════════════════════════════

Option Explicit

' ══════════ CONFIGURACIÓN ══════════
Const PALABRA_BUSCAR = "lista"
Const COLUMNA_clave = "clave"           ' Columna para deduplicar
Const EMAIL_DESTINATARIO = "sddsd@gmail.com"
Const ENVIAR_EMAIL = False
Const NOMBRE_SALIDA = "sin_horas"

' ══════════ COLUMNAS A MANTENER ══════════
' IMPORTANTE: Para agregar más columnas, añádelas aquí separadas por comas
' Ejemplo: "clave,nombre,apellido,telefono,email"
Const COLUMNAS_MANTENER = "clave,nombre,apellido"

' ══════════ VARIABLES ══════════
Dim fso, shell, carpetaDescargas, excelApp, wbPrincipal

Sub Main()
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    
    carpetaDescargas = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads\"
    If Not fso.FolderExists(carpetaDescargas) Then
        carpetaDescargas = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Descargas\"
    End If
    
    ' Buscar archivos
    Dim archivos()
    ReDim archivos(10)
    Dim contador, carpeta, archivo
    contador = 0
    
    Set carpeta = fso.GetFolder(carpetaDescargas)
    
    For Each archivo In carpeta.Files
        If InStr(LCase(archivo.Name), LCase(PALABRA_BUSCAR)) > 0 Then
            Dim rutaFinal
            If Right(LCase(archivo.Name), 4) <> ".xls" And Right(LCase(archivo.Name), 5) <> ".xlsx" Then
                rutaFinal = archivo.Path & ".xls"
                If Not fso.FileExists(rutaFinal) Then
                    On Error Resume Next
                    fso.CopyFile archivo.Path, rutaFinal
                    On Error GoTo 0
                End If
            Else
                rutaFinal = archivo.Path
            End If
            
            archivos(contador) = rutaFinal
            contador = contador + 1
        End If
    Next
    
    If contador < 2 Then
        MsgBox "❌ Se necesitan al menos 2 archivos que contengan '" & PALABRA_BUSCAR & "'" & vbCrLf & _
               "Solo se encontraron: " & contador, vbExclamation
        WScript.Quit
    End If
    
    ' Mostrar archivos
    Dim mensaje, i
    mensaje = "✓ Archivos encontrados:" & vbCrLf & vbCrLf
    For i = 0 To contador - 1
        mensaje = mensaje & (i + 1) & ". " & fso.GetFileName(archivos(i)) & vbCrLf
    Next
    mensaje = mensaje & vbCrLf & "Columnas a mantener: " & COLUMNAS_MANTENER & vbCrLf & vbCrLf
    mensaje = mensaje & "¿Continuar?"
    
    If MsgBox(mensaje, vbYesNo + vbQuestion) = vbNo Then
        WScript.Quit
    End If
    
    ' Iniciar Excel
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    
    Set wbPrincipal = excelApp.Workbooks.Add
    
    ' Importar archivos y filtrar columnas
    ImportarArchivo archivos(0), "Activos"
    FiltrarColumnas "Activos"
    
    ImportarArchivo archivos(1), "PorHoras"
    FiltrarColumnas "PorHoras"
    
    ' Procesar
    DeduplicarHoja "Activos"
    DeduplicarHoja "PorHoras"
    CruzarDatos
    
    ' Guardar
    Dim rutaSalida
    rutaSalida = carpetaDescargas & NOMBRE_SALIDA & ".xls"
    
    On Error Resume Next
    fso.DeleteFile rutaSalida
    wbPrincipal.SaveAs rutaSalida, 56
    On Error GoTo 0
    
    ' Limpiar
    wbPrincipal.Close False
    excelApp.Quit
    
    MsgBox "✓ ¡Completado!" & vbCrLf & vbCrLf & "Archivo guardado en:" & vbCrLf & rutaSalida, vbInformation
    
    If ENVIAR_EMAIL Then
        EnviarEmail rutaSalida
    End If
End Sub

Sub ImportarArchivo(ruta, nombreHoja)
    On Error Resume Next
    
    Dim wb, ws, wsNueva
    Set wb = excelApp.Workbooks.Open(ruta)
    
    If wb Is Nothing Then
        MsgBox "Error abriendo: " & ruta, vbCritical
        Exit Sub
    End If
    
    Set ws = wb.Worksheets(1)
    Set wsNueva = wbPrincipal.Worksheets.Add
    wsNueva.Name = nombreHoja
    
    ws.UsedRange.Copy wsNueva.Range("A1")
    wb.Close False
End Sub

Sub FiltrarColumnas(nombreHoja)
    ' Mantiene solo las columnas especificadas y elimina las demás
    On Error Resume Next
    
    Dim ws, i, j, columnaActual, nombreColumna
    Dim columnasAmantener, mantener, encontrada
    Dim ultimaColumna
    
    Set ws = wbPrincipal.Worksheets(nombreHoja)
    
    ' Dividir las columnas a mantener en un array
    columnasAmantener = Split(COLUMNAS_MANTENER, ",")
    
    ' Limpiar espacios
    For i = 0 To UBound(columnasAmantener)
        columnasAmantener(i) = Trim(columnasAmantener(i))
    Next
    
    ' Obtener última columna
    ultimaColumna = ws.Cells(1, ws.Columns.Count).End(-4159).Column
    
    ' Recorrer columnas de derecha a izquierda (para no desajustar índices al eliminar)
    For i = ultimaColumna To 1 Step -1
        nombreColumna = LCase(Trim(CStr(ws.Cells(1, i).Value)))
        
        ' Verificar si esta columna se debe mantener
        encontrada = False
        For j = 0 To UBound(columnasAmantener)
            If LCase(columnasAmantener(j)) = nombreColumna Then
                encontrada = True
                Exit For
            End If
        Next
        
        ' Si no está en la lista, eliminar la columna
        If Not encontrada Then
            ws.Columns(i).Delete
        End If
    Next
End Sub

Sub DeduplicarHoja(nombreHoja)
    On Error Resume Next
    
    Dim ws, colclave, i, dict, ultimaFila, fila, clave
    Set ws = wbPrincipal.Worksheets(nombreHoja)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Buscar columna clave
    colclave = 0
    For i = 1 To 50
        If LCase(Trim(CStr(ws.Cells(1, i).Value))) = LCase(COLUMNA_clave) Then
            colclave = i
            Exit For
        End If
    Next
    
    If colclave = 0 Then Exit Sub
    
    ultimaFila = ws.Cells(ws.Rows.Count, colclave).End(-4162).Row
    
    ' Eliminar duplicados por código
    For fila = ultimaFila To 2 Step -1
        clave = Trim(CStr(ws.Cells(fila, colclave).Value))
        If dict.Exists(clave) Or clave = "" Then
            ws.Rows(fila).Delete
        Else
            dict.Add clave, True
        End If
    Next
End Sub

Sub CruzarDatos()
    On Error Resume Next
    
    Dim wsActivos, wsPorHoras, wsResultado
    Dim dictHoras, colActivos, colPorHoras, i, clave, ultimaFila, filaDestino
    
    Set wsActivos = wbPrincipal.Worksheets("Activos")
    Set wsPorHoras = wbPrincipal.Worksheets("PorHoras")
    Set wsResultado = wbPrincipal.Worksheets.Add
    wsResultado.Name = "SinHoras"
    
    ' Buscar columna clave en ambas hojas
    colActivos = BuscarColumna(wsActivos)
    colPorHoras = BuscarColumna(wsPorHoras)
    
    If colActivos = 0 Or colPorHoras = 0 Then
        MsgBox "No se encontró columna '" & COLUMNA_clave & "'", vbExclamation
        Exit Sub
    End If
    
    ' Crear diccionario con los que TIENEN horas
    Set dictHoras = CreateObject("Scripting.Dictionary")
    ultimaFila = wsPorHoras.Cells(wsPorHoras.Rows.Count, colPorHoras).End(-4162).Row
    
    For i = 2 To ultimaFila
        clave = Trim(CStr(wsPorHoras.Cells(i, colPorHoras).Value))
        If clave <> "" Then
            dictHoras(clave) = True
        End If
    Next
    
    ' Copiar encabezados
    wsActivos.Rows(1).Copy wsResultado.Rows(1)
    
    ' Copiar solo los que NO tienen horas
    filaDestino = 2
    ultimaFila = wsActivos.Cells(wsActivos.Rows.Count, colActivos).End(-4162).Row
    
    For i = 2 To ultimaFila
        clave = Trim(CStr(wsActivos.Cells(i, colActivos).Value))
        
        If clave <> "" And Not dictHoras.Exists(clave) Then
            wsActivos.Rows(i).Copy wsResultado.Rows(filaDestino)
            wsResultado.Rows(filaDestino).Interior.Color = RGB(255, 199, 206)
            filaDestino = filaDestino + 1
        End If
    Next
    
    ' Formato
    wsResultado.Rows(1).Font.Bold = True
    wsResultado.Rows(1).Interior.Color = RGB(192, 0, 0)
    wsResultado.Rows(1).Font.Color = RGB(255, 255, 255)
    wsResultado.Columns.AutoFit
End Sub

Function BuscarColumna(ws)
    Dim i
    For i = 1 To 50
        If LCase(Trim(CStr(ws.Cells(1, i).Value))) = LCase(COLUMNA_clave) Then
            BuscarColumna = i
            Exit Function
        End If
    Next
    BuscarColumna = 0
End Function

Sub EnviarEmail(rutaArchivo)
    On Error Resume Next
    Dim outlook, mail
    Set outlook = CreateObject("Outlook.Application")
    Set mail = outlook.CreateItem(0)
    mail.To = EMAIL_DESTINATARIO
    mail.Subject = "Informe Sin Horas - Cruz Roja"
    mail.Body = "Adjunto informe automático."
    mail.Attachments.Add rutaArchivo
    mail.Display
End Sub

Main
