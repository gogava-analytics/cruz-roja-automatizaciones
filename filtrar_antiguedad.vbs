' ═══════════════════════════════════════════════════════════════════════════
' AUTOMATIZACIÓN CRUZ ROJA - PROCESAR GETJOBID
' ═══════════════════════════════════════════════════════════════════════════

Option Explicit

' ══════════════════════════════════════════════════════════════
'  CONFIGURACION — solo toca esto si algo cambia
' ══════════════════════════════════════════════════════════════
Const PALABRA_BUSCAR      = "getjobid"
Const COLUMNA_FECHA       = "FECHA_INICIO"
Const DIAS_MINIMOS        = 7
Const NOMBRE_SALIDA       = "pasadas_una_semana"
Const COLUMNA_DIAS_NOMBRE = "DIAS_TRANSCURRIDOS"

' Columnas que se mantienen en el resultado (separadas por coma)
Const COLUMNAS_MANTENER   = "NOMBRE1,CF_LIT_CENTR,LITERAL,FECHA_INICIO,CODIGO,NOMBRE"

' ══════════ EMAIL (opcional) ══════════
Const ENVIAR_EMAIL        = False
Const EMAIL_DESTINATARIO  = "sddsd@gmail.com"

' ══════════════════════════════════════════════════════════════
'  VARIABLES GLOBALES
' ══════════════════════════════════════════════════════════════
Dim fso, shell, carpetaDescargas, excelApp

' ══════════════════════════════════════════════════════════════
'  INICIO
' ══════════════════════════════════════════════════════════════
Sub Main()
    Set fso   = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    carpetaDescargas = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads\"
    If Not fso.FolderExists(carpetaDescargas) Then
        carpetaDescargas = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Descargas\"
    End If
    If Not fso.FolderExists(carpetaDescargas) Then
        MsgBox "No se encontro la carpeta Descargas.", vbCritical
        WScript.Quit
    End If

    Dim rutaArchivo
    rutaArchivo = BuscarArchivo(carpetaDescargas, PALABRA_BUSCAR)
    If rutaArchivo = "" Then
        MsgBox "No se encontro ningun archivo con '" & PALABRA_BUSCAR & "' en:" & vbCrLf & carpetaDescargas, vbExclamation
        WScript.Quit
    End If

    Dim rutaXls
    rutaXls = AsegurarExtensionXls(rutaArchivo)

    If MsgBox("Archivo: " & fso.GetFileName(rutaXls) & vbCrLf & vbCrLf & "Continuar?", vbYesNo + vbQuestion, "Procesar GetJobId") = vbNo Then
        WScript.Quit
    End If

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible       = False
    excelApp.DisplayAlerts = False

    ProcesarArchivo rutaXls

    If ENVIAR_EMAIL Then
        EnviarEmail carpetaDescargas & NOMBRE_SALIDA & ".xls"
    End If
End Sub


' ══════════════════════════════════════════════════════════════
'  BUSCAR ARCHIVO
' ══════════════════════════════════════════════════════════════
Function BuscarArchivo(carpeta, palabra)
    Dim objCarpeta, archivo
    Set objCarpeta = fso.GetFolder(carpeta)
    BuscarArchivo = ""
    For Each archivo In objCarpeta.Files
        If InStr(LCase(archivo.Name), LCase(palabra)) > 0 Then
            BuscarArchivo = archivo.Path
            Exit Function
        End If
    Next
End Function


' ══════════════════════════════════════════════════════════════
'  AÑADIR .XLS SI NO TIENE EXTENSION EXCEL
' ══════════════════════════════════════════════════════════════
Function AsegurarExtensionXls(ruta)
    Dim ext
    ext = LCase(Right(ruta, 4))
    If ext = ".xls" Or ext = "xlsx" Then
        AsegurarExtensionXls = ruta
        Exit Function
    End If
    Dim nuevaRuta
    nuevaRuta = ruta & ".xls"
    If Not fso.FileExists(nuevaRuta) Then
        On Error Resume Next
        fso.CopyFile ruta, nuevaRuta
        On Error GoTo 0
    End If
    AsegurarExtensionXls = nuevaRuta
End Function


' ══════════════════════════════════════════════════════════════
'  PROCESAR: filtra fechas y mantiene solo columnas elegidas
' ══════════════════════════════════════════════════════════════
Sub ProcesarArchivo(rutaXls)

    ' Abrir origen
    Dim wbOrigen, wsOrigen
    On Error Resume Next
    Set wbOrigen = excelApp.Workbooks.Open(rutaXls)
    On Error GoTo 0
    If wbOrigen Is Nothing Then
        MsgBox "No se pudo abrir: " & rutaXls, vbCritical
        excelApp.Quit
        WScript.Quit
    End If
    Set wsOrigen = wbOrigen.Worksheets(1)

    ' Encontrar columna FECHA_INICIO
    Dim colFecha
    colFecha = BuscarColumna(wsOrigen, COLUMNA_FECHA)
    If colFecha = 0 Then
        MsgBox "No se encontro la columna '" & COLUMNA_FECHA & "'.", vbExclamation
        wbOrigen.Close False
        excelApp.Quit
        WScript.Quit
    End If

    ' Leer lista de columnas a mantener
    Dim listaColumnas
    listaColumnas = Split(COLUMNAS_MANTENER, ",")
    Dim i
    For i = 0 To UBound(listaColumnas)
        listaColumnas(i) = Trim(listaColumnas(i))
    Next

    ' Buscar en qué numero de columna esta cada una en el origen
    Dim indicesOrigen()
    ReDim indicesOrigen(UBound(listaColumnas))
    For i = 0 To UBound(listaColumnas)
        indicesOrigen(i) = BuscarColumna(wsOrigen, listaColumnas(i))
    Next

    ' Crear libro de salida
    Dim wbNuevo, wsNuevo
    Set wbNuevo = excelApp.Workbooks.Add
    Set wsNuevo = wbNuevo.Worksheets(1)
    wsNuevo.Name = "Resultado"

    ' Escribir encabezados (solo columnas elegidas + DIAS_TRANSCURRIDOS al final)
    For i = 0 To UBound(listaColumnas)
        wsNuevo.Cells(1, i + 1).Value = listaColumnas(i)
    Next
    Dim colDias
    colDias = UBound(listaColumnas) + 2
    wsNuevo.Cells(1, colDias).Value = COLUMNA_DIAS_NOMBRE

    ' Formato encabezado: negrita y fondo gris simple
    Dim totalCols
    totalCols = colDias
    wsNuevo.Range(wsNuevo.Cells(1, 1), wsNuevo.Cells(1, totalCols)).Font.Bold = True
    wsNuevo.Range(wsNuevo.Cells(1, 1), wsNuevo.Cells(1, totalCols)).Interior.Color = RGB(217, 217, 217)

    ' Filtrar filas por fecha y copiar solo columnas elegidas
    Dim ultimaFila
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, colFecha).End(-4162).Row

    Dim filaDestino
    filaDestino = 2

    Dim hoy
    hoy = Date

    Dim j
    For j = 2 To ultimaFila
        Dim valorFecha
        valorFecha = wsOrigen.Cells(j, colFecha).Value

        Dim esValida
        esValida = False
        If valorFecha <> "" Then
            If IsDate(valorFecha) Then
                esValida = True
            End If
        End If

        If esValida Then
            Dim diasTranscurridos
            diasTranscurridos = DateDiff("d", CDate(valorFecha), hoy)

            If diasTranscurridos >= DIAS_MINIMOS Then
                ' Copiar solo las columnas elegidas
                For i = 0 To UBound(listaColumnas)
                    If indicesOrigen(i) > 0 Then
                        wsNuevo.Cells(filaDestino, i + 1).Value = wsOrigen.Cells(j, indicesOrigen(i)).Value
                    End If
                Next
                ' Escribir dias transcurridos
                wsNuevo.Cells(filaDestino, colDias).Value = diasTranscurridos
                filaDestino = filaDestino + 1
            End If
        End If
    Next

    ' Autoajustar columnas
    wsNuevo.Columns.AutoFit

    ' Guardar
    Dim rutaSalida
    rutaSalida = carpetaDescargas & NOMBRE_SALIDA & ".xls"
    If fso.FileExists(rutaSalida) Then
        On Error Resume Next
        fso.DeleteFile rutaSalida
        On Error GoTo 0
    End If
    wbNuevo.SaveAs rutaSalida, 56
    wbNuevo.Close False
    wbOrigen.Close False
    excelApp.Quit

    MsgBox "Completado!" & vbCrLf & vbCrLf & _
           "Filas con mas de " & DIAS_MINIMOS & " dias: " & (filaDestino - 2) & vbCrLf & vbCrLf & _
           "Guardado en: " & rutaSalida, vbInformation
End Sub


' ══════════════════════════════════════════════════════════════
'  BUSCAR NUMERO DE COLUMNA POR NOMBRE
' ══════════════════════════════════════════════════════════════
Function BuscarColumna(ws, nombreColumna)
    Dim i
    BuscarColumna = 0
    For i = 1 To 100
        If LCase(Trim(CStr(ws.Cells(1, i).Value))) = LCase(Trim(nombreColumna)) Then
            BuscarColumna = i
            Exit Function
        End If
    Next
End Function


' ══════════════════════════════════════════════════════════════
'  ENVIAR EMAIL — eliminar o dejar ENVIAR_EMAIL = False si no se usa
' ══════════════════════════════════════════════════════════════
Sub EnviarEmail(rutaArchivo)
    On Error Resume Next
    Dim outlook, mail
    Set outlook      = CreateObject("Outlook.Application")
    Set mail         = outlook.CreateItem(0)
    mail.To          = EMAIL_DESTINATARIO
    mail.Subject     = "Informe GetJobId - Cruz Roja"
    mail.Body        = "Adjunto informe de registros con mas de " & DIAS_MINIMOS & " dias."
    mail.Attachments.Add rutaArchivo
    mail.Display
End Sub


Main
