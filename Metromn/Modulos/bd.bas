Attribute VB_Name = "BD"
' Public conn As ADODB.Connection
Public Function datos_bd(ByVal consulta As String) As ADODB.Recordset
    On Error GoTo fallo
    log (consulta)
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn) = True Then ' CONECTAR EL RS
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
    msj = "Error en el acceso a la bd: " & Err.Description
'    MsgBox msj, vbCritical, App.Title
    log msj, 1
End Function

Public Function CrearConexionGlobal(conn As ADODB.Connection) As Boolean
    On Error GoTo falloConexion
        Set conn = New ADODB.Connection
        conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" _
                              & "SERVER=" & frmInformes.txtmysql(0) & ";" _
                              & "DATABASE=" & frmInformes.txtmysql(1) & ";" _
                              & "UID=" & frmInformes.txtmysql(2) & ";" _
                              & "PWD=" & frmInformes.txtmysql(3) & ";" _
                              & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
        conn.Open
        frmInformes.chkConectado.Value = 1
    CrearConexionGlobal = True
    Exit Function
falloConexion:
    CrearConexionGlobal = False
    MsgBox "Error en la conexion :" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
Public Sub log(datos As String, Optional error As Integer)
    On Error Resume Next
    If frmVersion.chklog = unchecked Then Exit Sub
    On Error GoTo fallo
    Open App.Path & "\log.txt" For Append As #1
    Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
    Close
    Exit Sub
fallo:
    Close
    Exit Sub
End Sub
Public Function execute_bd(ByVal consulta As String) As ADODB.Recordset
    log (consulta)
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn) = True Then ' CONECTAR EL RS
        conn.Execute consulta
    End If
    Set conn = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
'    MsgBox msj, vbCritical, App.Title
    log msj, 1
End Function
Public Sub msg_error(Mensaje As String)
    MsgBox Mensaje, vbCritical, App.Title
End Sub
Public Function Eliminar_Caracteres(cadena As String)
    Dim nombre As String
    nombre = cadena
    nombre = Replace(nombre, """", "-")
    nombre = Replace(nombre, "\", "")
    nombre = Replace(nombre, "'", "")
    nombre = Replace(nombre, "%", "")
    nombre = Replace(nombre, "_", "")
    nombre = Replace(nombre, "<", "")
    nombre = Replace(nombre, ">", "")
    Eliminar_Caracteres = nombre
End Function
