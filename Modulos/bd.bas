Attribute VB_Name = "bd"
'Public conn As ADODB.Connection 'conexion global
Public Function datos_bd(ByVal consulta As String, Optional no_log As Boolean) As ADODB.Recordset
    On Error GoTo fallo
    If Not no_log Then
        log (consulta)
    End If
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        Dim rs As New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        rs.Open consulta
        Set rs.ActiveConnection = Nothing ' DESCONECTAR EL RS
        Set datos_bd = rs
        Set rs = Nothing
        conn.Close
    Else
        MsgBox "Error al crear la conexión.", vbCritical, App.Title
    End If
    Set conn = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description & " : " & Err.Source & " : " & Err.Number
    error_grave_jgm msj
    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function execute_bd(ByVal consulta As String, Optional no_actualizar As Boolean) As ADODB.Recordset
    If Not no_actualizar Then
        log (consulta)
    End If
    
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        conn.Execute SQL_BT & consulta & SQL_ROLL
        If Not no_actualizar Then
            Dim rs As New ADODB.Recordset
            
            rs.ActiveConnection = conn
            rs.CursorLocation = adUseClient
            rs.CursorType = adOpenForwardOnly
            rs.LockType = adLockReadOnly
            rs.Open "SELECT ROW_COUNT();"
            Dim reg As Integer
            If rs.RecordCount > 0 Then
                reg = rs(0)
            Else
                reg = 0
            End If
            Set rs.ActiveConnection = Nothing ' DESCONECTAR EL RS
            Set rs = Nothing
'            Set rs = datos_bd("SELECT ROW_COUNT();")
'            Dim reg As Integer
'            If rs.RecordCount > 0 Then
'                reg = rs(0)
'            Else
'                reg = 0
'            End If
            insertar_actualizaciones consulta, reg
        End If
    Else
        MsgBox "Error al crear la conexión.", vbCritical, App.Title
    End If
    Set conn = Nothing
    
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
    error_grave_jgm msj
    log (msj)
    MsgBox msj, vbCritical, App.Title
End Function
Public Function CrearConexionGlobal(conn As ADODB.Connection, user As String, PassWord As String) As Boolean
    Dim ipRegistro As String
    Dim database As String
    Dim driver As String
    Dim port As String
    On Error GoTo falloConexion
    Set conn = New ADODB.Connection
'    If UCase(user) = "PRUEBA" Then
    If MODO_PRUEBA Then
'       database = ReadINI(App.Path + "\config.ini", "server", "bd_prueba")
        ipRegistro = ReadINI(App.Path + "\config.ini", "server_prueba", "ip")
        database = ReadINI(App.Path + "\config.ini", "server_prueba", "bd")
        driver = ReadINI(App.Path + "\config.ini", "server_prueba", "DRIVER")
        port = ReadINI(App.Path + "\config.ini", "server_prueba", "PORT")
    Else
        ipRegistro = ReadINI(App.Path + "\config.ini", "server", "ip")
        database = ReadINI(App.Path + "\config.ini", "server", "bd")
        driver = ReadINI(App.Path + "\config.ini", "server", "DRIVER")
        port = ReadINI(App.Path + "\config.ini", "server", "PORT")
    End If
    If port = "" Then
        port = "3306"
    End If
    conn.ConnectionString = "DRIVER=" & driver & ";" _
                            & "SERVER=" & ipRegistro & ";" _
                            & "PORT=" & port & ";" _
                            & "DATABASE=" & database & ";" _
                            & "UID=" & BD_USUARIO & ";" _
                            & "PWD=" & BD_PASS & ";" _
                            & "OPTION=16427"
    conn.Open
    CrearConexionGlobal = True
    Exit Function
falloConexion:
    CrearConexionGlobal = False
    MsgBox "CrearConexionGlobal:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
Public Sub llenar_combo(combo As miCombo, oB As Object, FK As Long, FORMULARIO As Form, filtro As String)
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        oB.llenar_combo conn, combo, FK, FORMULARIO, filtro
    End If
    Set conn = Nothing
End Sub
Public Function insertar_actualizaciones(ByVal consulta As String, Optional registros As Integer) As Boolean
    Dim c As String
   On Error GoTo insertar_actualizaciones_Error

'    If UCase(usuario.getNOMBRE) = "PRUEBA" Then
    If MODO_PRUEBA Then
'       database = ReadINI(App.Path + "\config.ini", "server", "bd_prueba")
       database = ReadINI(App.Path + "\config.ini", "server_prueba", "bd")
    Else
       database = ReadINI(App.Path + "\config.ini", "server", "bd")
    End If
    database = database & "_actualizaciones"
    
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        c = "INSERT INTO " & database & ".actualizaciones (FTIMESTP,CONSULTA,REGISTROS,ACTUALIZADA) VALUES (CURRENT_TIMESTAMP,'" & UCase(Trim(Replace(consulta, "'", "#"))) & "'," & registros & "," & USUARIO.getID_EMPLEADO & ")"
        conn.Execute c
    End If
    Set conn = Nothing
    
   On Error GoTo 0
   Exit Function

insertar_actualizaciones_Error:
    Dim msj As String
    msj = "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_actualizaciones of Módulo bd"
    error_grave_jgm msj
    MsgBox msj
End Function

