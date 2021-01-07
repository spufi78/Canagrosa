Attribute VB_Name = "general"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal Hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public USUARIO As New clsUsuarios
Public conn As ADODB.Connection
Public database As String
Public referencia_word As String
Public referencia_pdf As String

Public DIRECTORIO_TEMPORAL As String

Public Function datos_bd(ByVal consulta As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    log (consulta)
    rs.Open consulta
    Set datos_bd = rs
    Set rs = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description
'    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function execute_bd(ByVal consulta As String, Optional Actualizar As Boolean) As ADODB.Recordset
    log (consulta)
    conn.Execute consulta
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
'    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function CrearConexionGlobal() As Boolean
    Dim ipRegistro As String
    On Error GoTo falloConexion
    ipRegistro = ReadINI(App.Path + "\config.ini", "server", "ip")
    database = ReadINI(App.Path + "\config.ini", "server", "bd")
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DRIVER=" & ReadINI(App.Path + "\config.ini", "SERVER", "DRIVER") & ";" _
                            & "SERVER=" & ipRegistro & ";" _
                            & "DATABASE=" & database & ";" _
                            & "UID=geslab_huesna;" _
                            & "PWD=ix1tec;" _
                            & "OPTION=3"
'                            & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    conn.Open
    CrearConexionGlobal = True
    Exit Function
falloConexion:
    CrearConexionGlobal = False
    MsgBox "CrearConexionGlobal:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
Public Sub Espera(Segundos As Single)
  Dim ComienzoSeg As Single
  Dim FinSeg As Single
  ComienzoSeg = Timer
  FinSeg = ComienzoSeg + Segundos
  Do While FinSeg > Timer
      DoEvents
      If ComienzoSeg > Timer Then
          FinSeg = FinSeg - 24 * 60 * 60
      End If
  Loop
End Sub

Function Encripta(Strg$, PassWord$)
   Dim b$, s$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, p$
   j = 1
   For i = 1 To Len(PassWord$)
     p$ = p$ & Asc(Mid$(PassWord$, i, 1))
   Next
    
   For i = 1 To Len(Strg$)
     A1 = Asc(Mid$(p$, j, 1))
     j = j + 1: If j > Len(p$) Then j = 1
     A2 = Asc(Mid$(Strg$, i, 1))
     A3 = A1 Xor A2
     b$ = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b$
     s$ = s$ + b$
   Next
   Encripta = s$
End Function
Public Sub log(datos As String)
    On Error Resume Next
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log"
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date)
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date) & "\pdf"
    On Error GoTo fallo
    Open ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date) & "\pdf\" & Format(Date, "yyyy-mm-dd") & " PDF.txt" For Append As #1
    If Left(datos, 36) <> "SELECT A.MUESTRA_ID,A.TIPO,E.USUARIO" Then
        Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
    End If
    Close
    Exit Sub
fallo:
    Close
    Exit Sub
End Sub
Public Function subindices(texto As String) As Boolean
    subindices = False
    If InStr(1, UCase(texto), "SUP(", vbTextCompare) > 0 Or _
       InStr(1, UCase(texto), "SUB(", vbTextCompare) > 0 Then
        subindices = True
    End If
End Function
'Public Sub texto_formateado(texto As String, rango As Range, parentesis As Integer)
'    If subindices(texto) = True Then
'       If parentesis = 1 Then
'        rango.InsertAfter " ("
'       End If
'       Dim activo As Boolean
'       activo = False
'       h = 1
'       While h <= Len(texto)
'           If UCase(Mid(texto, h, 3)) = "SUP" Or _
'              UCase(Mid(texto, h, 3)) = "SUB" Then
'                pos = InStr(h, texto, ")")
'                rango.InsertAfter Mid(texto, h + 4, pos - (h + 4))
'                If UCase(Mid(texto, h, 3)) = "SUP" Then
'                    For i = (pos - (h + 4)) To 1 Step -1
'                        rango.Characters(Len(rango.Text) - (i + 1)).Font.Superscript = True
'                    Next
'                Else
'                    For i = (pos - (h + 4)) To 1 Step -1
'                        rango.Characters(Len(rango.Text) - (i + 1)).Font.Subscript = True
'                    Next
'                End If
'                activo = True
'                h = pos
'           Else
''                Debug.Print Mid(texto, h, 1) & " -> " & Asc(Mid(texto, h, 1))
'                If Asc(Mid(texto, h, 1)) <> 10 Then
''                    Select Case Asc(Mid(texto, h, 1))
''                    Case 10
''                        rango.InsertAfter vbNewLine
''                    Case Else
'                        rango.InsertAfter Mid(texto, h, 1)
'                        If activo = True Then
'                            rango.Characters(Len(rango.Text) - 2).Font.Superscript = False
'                            rango.Characters(Len(rango.Text) - 2).Font.Subscript = False
'                            activo = False
'                        End If
''                    End Select
'                End If
'           End If
'           h = h + 1
'       Wend
'       If parentesis = 1 Then
'        rango.InsertAfter ")"
'        rango.Characters(Len(rango.Text) - 2).Font.Superscript = False
'        rango.Characters(Len(rango.Text) - 2).Font.Subscript = False
'       End If
'    Else
'       If parentesis = 1 And Trim(texto) <> "" Then
'           rango.InsertAfter " (" & texto & ")"
'       Else
'        rango.InsertAfter texto
'       End If
'    End If
'End Sub

Public Function fecha_bd(ByVal fecha As String) As String
If Trim(fecha) = "" Then fecha = "0000-00-00"
fecha = Format(fecha, "yyyy-mm-dd")

fecha_bd = fecha

End Function
Function moneda_bd(valor As String) As String
    valor = IIf(Trim(valor) = "", "0", valor)
    
    If UCase(ReadINI(App.Path + "\config.ini", "server", "tipo")) = "ACCESS" Then
        moneda_bd = Format(valor, "currency")
    Else
        moneda_bd = Replace(Format(valor, "0.00"), ",", ".")
    End If
End Function

Public Function Eliminar_Caracteres_Archivo(cadena As String)
    Dim nombreNuevo As String
    nombreNuevo = cadena
    nombreNuevo = Replace(nombreNuevo, ":", "")
    nombreNuevo = Replace(nombreNuevo, "/", "")
    nombreNuevo = Replace(nombreNuevo, "\", "")
    nombreNuevo = Replace(nombreNuevo, "*", "")
    nombreNuevo = Replace(nombreNuevo, "?", "")
    nombreNuevo = Replace(nombreNuevo, "<", "")
    nombreNuevo = Replace(nombreNuevo, ">", "")
    nombreNuevo = Replace(nombreNuevo, "'", "")
    Eliminar_Caracteres_Archivo = nombreNuevo
End Function


Public Sub enviar_informe_error(MUESTRA As Long, ERROR As String)
    On Error Resume Next
    Dim sPara As String
    Dim sAsunto As String
    Dim sMensaje As String
    Dim sFichero_Log As String
    sPara = "julio.gonzalez@ixitec.net"
    sAsunto = "Error al generar la muestra (ID : " & MUESTRA & ")"
    sMensaje = sMensaje & vbNewLine & "*****************************"
    sMensaje = sMensaje & vbNewLine & " ID : " & MUESTRA
    sMensaje = sMensaje & vbNewLine & " FECHA : " & Date
    sMensaje = sMensaje & vbNewLine & " HORA : " & Time
    sMensaje = sMensaje & vbNewLine & " ERROR : " & ERROR
    sMensaje = sMensaje & vbNewLine & "*****************************"
    sFichero_Log = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date) & "\pdf\" & Format(Date, "yyyy-mm-dd") & " PDF.txt"
    Enviar_Mail_CDO sPara, sAsunto, sMensaje, sFichero_Log
End Sub
Public Function insertar_actualizaciones(ByVal consulta As String, Optional registros As Integer) As Boolean
    On Error Resume Next
    Dim c As String
    c = "INSERT INTO geslab_huesna_actualizaciones.ACTUALIZACIONES (FTIMESTP,CONSULTA,REGISTROS,ACTUALIZADA) VALUES (CURRENT_TIMESTAMP,'" & UCase(Trim(Replace(consulta, "'", "#"))) & "'," & registros & "," & USUARIO.getID_EMPLEADO & ")"
    conn.Execute c
End Function
