Attribute VB_Name = "funciones"
Public conn As ADODB.Connection
Public Function datos_bd(ByVal consulta As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim msj As String
    
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    rs.Open consulta
    Set datos_bd = rs
    Set rs = Nothing
    Exit Function

fallo:
    
    msj = "Error en el acceso a la bd: " & Err.Description & " : " & Err.Source & " : " & Err.Number
    MsgBox msj, vbCritical, App.Title
    
End Function
Public Function execute_bd(ByVal consulta As String, Optional no_actualizar As Boolean) As ADODB.Recordset
    
    Dim msj As String
    
    conn.Execute SQL_BT & consulta & SQL_ROLL
    If Not no_actualizar Then
        insertar_actualizaciones (consulta)
    End If
    
    Exit Function
fallo:
    
    msj = "Error en el execute bd: " & Err.Description
    MsgBox msj, vbCritical, App.Title
    
End Function
Public Function insertar_actualizaciones(ByVal consulta As String) As Boolean
    On Error Resume Next
    Dim c As String
    c = "INSERT INTO ACTUALIZACIONES (FTIMESTP,CONSULTA) VALUES (CURRENT_TIMESTAMP,'" & UCase(Trim(Replace(consulta, "'", "#"))) & "')"
    conn.Execute c
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



