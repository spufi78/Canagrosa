Attribute VB_Name = "bd_desconectada"
Public conn As ADODB.Connection
Public Function datos_bd(ByVal CONSULTA As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    rs.Open CONSULTA
    Set datos_bd = rs
    Set rs = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description
    MsgBox msj, vbCritical, App.Title
End Function
Public Function execute_bd(ByVal CONSULTA As String) As ADODB.Recordset
    conn.Execute CONSULTA
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
    MsgBox msj, vbCritical, App.Title
End Function
Public Function CrearConexion() As ADODB.Connection
    Dim ipRegistro As String
    Dim database As String
    On Error GoTo falloConexion
'    Set conn = New ADODB.Connection
'    ipRegistro = ReadINI(App.Path + "\config.ini", "server", "ip")
'    If UCase(user) = "PRUEBA" Then
'       database = ReadINI(App.Path + "\config.ini", "server", "bd_prueba")
'    Else
'       database = ReadINI(App.Path + "\config.ini", "server", "bd")
'    End If
'    conn.ConnectionString = "DRIVER=" & ReadINI(App.Path + "\config.ini", "SERVER", "DRIVER") & ";" _
'                            & "SERVER=" & ipRegistro & ";" _
'                            & "DATABASE=" & database & ";" _
'                            & "UID=" & ReadINI(App.Path + "\config.ini", "SERVER", "BD_USUARIO") & ";" _
'                            & "PWD=" & ReadINI(App.Path + "\config.ini", "SERVER", "BD_PASS") & ";" _
'                            & "OPTION=16427"
    
'    frmBusquedaGeneral.lblip = conn.ConnectionString
'    conn.Open
    Exit Function
falloConexion:
'    CrearConexion = vbNull
    MsgBox "CrearConexionGlobal:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
Public Sub cargar_combo(combo As DataCombo, PK As String, CAMPO As String, TABLA As String, FILTRO As String, QUERY As String)
    Dim rs As ADODB.Recordset
    Dim CONSULTA As String
    Dim s As String
    If QUERY <> "" Then
        CONSULTA = QUERY & " ORDER BY " & CAMPO
    Else
        If FILTRO <> "" Then
            s = " WHERE " & FILTRO & " "
        End If
        If PK <> "" And CAMPO <> "" And TABLA <> "" Then
            CONSULTA = "SELECT " & PK & "," & CAMPO & _
                       "  FROM " & TABLA & _
                       s & _
                       " ORDER BY " & CAMPO
        End If
    End If
    Set rs = datos_bd(CONSULTA)
    Set combo.RowSource = rs
    combo.ListField = rs(1).Name
    combo.BoundColumn = rs(0).Name
    Set rs = Nothing
End Sub
Public Sub cargar_combo_FK(combo As DataCombo, PK As String, CAMPO As String, TABLA As String, FK_CAMPO As String, FK_VALOR As Long, FILTRO As String, QUERY As String)
    Dim rs As ADODB.Recordset
    Dim CONSULTA As String
    Dim s As String
    If QUERY <> "" Then
        CONSULTA = QUERY
    Else
        If FILTRO <> "" Then
            s = " AND " & FILTRO & " "
        End If
        CONSULTA = "SELECT " & PK & "," & CAMPO & _
                   "  FROM " & TABLA & _
                   " WHERE " & FK_CAMPO & " = " & FK_VALOR & _
                   s & _
                   " ORDER BY " & CAMPO
    End If
    Set rs = datos_bd(CONSULTA)
    Set combo.RowSource = rs
    combo.ListField = rs(1).Name
    combo.BoundColumn = rs(0).Name
    Set rs = Nothing
End Sub


