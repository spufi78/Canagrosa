Attribute VB_Name = "bd"
Public conn As ADODB.Connection
Public cConnContabilidad As ADODB.Connection

Public Function datos_bd(ByVal consulta As String) As ADODB.Recordset
    On Error GoTo fallo
'    log (consulta)
    Dim rs As New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    rs.Open consulta
    Set datos_bd = rs
    Set rs = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description
    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function

Public Function CrearConexionGlobal(USUARIO As String, user As String, pass As String, prueba As Boolean) As Boolean
    Dim ipRegistro As String
    Dim database As String
    On Error GoTo falloConexion
    If Left(USUARIO, 1) = "-" Then
        ipRegistro = IP_RESPALDO
    Else
        ipRegistro = ReadINI(App.Path + "\config.ini", "server", "ip")
    End If
    ip = ipRegistro
    Set conn = New ADODB.Connection
    If UCase(ReadINI(App.Path + "\config.ini", "server", "tipo")) = "ACCESS" Then
       Dim ruta As String
       ruta = ReadINI(App.Path + "\config.ini", "server", "ruta")
       conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ruta & "\" & ReadINI(App.Path + "\config.ini", "server", "bd") & ";" ' Jet OLEDB:Database Password=" & pass
    Else
        If prueba Then
           database = ReadINI(App.Path + "\config.ini", "server", "bd_prueba")
        Else
           database = ReadINI(App.Path + "\config.ini", "server", "bd")
        End If
        gbd = database
       conn.ConnectionString = "DRIVER=" & ReadINI(App.Path + "\config.ini", "SERVER", "DRIVER") & ";" _
                             & "SERVER=" & ipRegistro & ";" _
                             & "DATABASE=" & database & ";" _
                             & "UID=" & user & ";" _
                             & "PWD=" & pass & ";" _
                             & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    End If
    conn.Open
    CrearConexionGlobal = True
    Exit Function
falloConexion:
    CrearConexionGlobal = False
    MsgBox "CrearConexionGlobal:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
Public Sub log(datos As String)
    On Error Resume Next
    If datos = "select current_timestamp" Then
        Exit Sub
    End If
    If USUARIO.getUSUARIO <> "" Then
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Log") & "\" & Year(Date)
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Log") & "\" & Year(Date) & "\" & Format(Date, "mmmm")
        On Error GoTo fallo
        Open ReadINI(App.Path + "\config.ini", "Documentos", "Log") & "\" & Year(Date) & "\" & Format(Date, "mmmm") & "\" & Format(Date, "yyyy-mm-dd") & " " & UCase(USUARIO.getUSUARIO) & ".txt" For Append As #1
        If Left(datos, 3) = "frm" Then
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & String(75, "-")
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & String(75, "-")
        Else
            If InStr(1, datos, "Error en el acceso a la bd") > 0 Or _
               InStr(1, datos, "Error en el execute bd") > 0 Then
                Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & String(Len(datos), "*")
                Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
                Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & String(Len(datos), "*")
            Else
                Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
            End If
        End If
    End If
    Close
    Exit Sub
fallo:
    Close
    Exit Sub
End Sub
Public Function execute_bd(ByVal consulta As String) As ADODB.Recordset
'    log (consulta)
    On Error GoTo fallo
    conn.Execute Trim(Replace(consulta, vbCrLf, " -> "))
'    conn.Execute consulta & ";"
'    insertar_actualizaciones (consulta)
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function insertar_actualizaciones(ByVal consulta As String) As Boolean
    On Error Resume Next
    Dim c As String
    c = "INSERT INTO ACTUALIZACIONES (FTIMESTP,CONSULTA,ACTUALIZADA) VALUES (CURRENT_TIMESTAMP,'" & UCase(Trim(Replace(consulta, "'", "#"))) & "'," & USUARIO.getID_EMPLEADO & ")"
    conn.Execute c
End Function
Public Sub msg_error(Mensaje As String)
    MsgBox Mensaje, vbCritical, App.Title
End Sub
Public Sub Cargar_Combo(combo As DataCombo, ob As Object)
    Dim rs As ADODB.Recordset
    Set rs = ob.Listado_Combo
    Set combo.RowSource = rs
    combo.ListField = rs(1).Name
    combo.BoundColumn = rs(0).Name
    Set rs = Nothing
End Sub
Public Sub cargar_combo_FK(combo As DataCombo, ob As Object, pk As Long)
    Dim rs As ADODB.Recordset
    Set rs = ob.Listado_Combo_FK(pk)
    Set combo.RowSource = rs
    combo.ListField = rs(1).Name
    combo.BoundColumn = rs(0).Name
    Set rs = Nothing
End Sub

Public Sub Posicionar(lista As ListView, Col As Integer, texto As String)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If Col = 0 Then
            If lista.ListItems(i) = texto Then
                lista.ListItems(i).Selected = True
                lista.ListItems(i).EnsureVisible
                lista.SetFocus
            End If
        Else
            If lista.ListItems(i).SubItems(Col) = texto Then
                lista.ListItems(i).Selected = True
                lista.ListItems(i).EnsureVisible
                lista.SetFocus
            End If
        End If
    Next
End Sub
