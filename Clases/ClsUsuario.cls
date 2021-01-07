VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private permisos(5) As Boolean
Private datosUsuario(3) As String
Private USO As String
Private estado1 As Boolean
Public Property Get ESTADO() As Boolean
    ESTADO = estado1
End Property
Public Property Get id() As Integer
    id = val(datosUsuario(0))
End Property
Public Property Get NOMBRE() As String
    NOMBRE = datosUsuario(1)
End Property
Public Property Get APELLIDOS() As String
    APELLIDOS = datosUsuario(2)
End Property
Public Property Get USUARIO() As String
    USUARIO = datosUsuario(3)
End Property
Public Property Get getUSO() As String
    getUSO = USO
End Property
Public Property Let setUSO(ByVal dato As String)
    USO = Trim(dato)
End Property
Public Property Get PermisoImpresion() As Boolean
    PermisoImpresion = permisos(0)
End Property
Public Property Get Permisofacturacion() As Boolean
    Permisofacturacion = permisos(1)
End Property
Public Property Get PermisoModificacion() As Boolean
    PermisoModificacion = permisos(2)
End Property
Public Property Get PermisoEliminiacion() As Boolean
    PermisoEliminiacion = permisos(3)
End Property
Public Property Get PermisoCrearUsuario() As Boolean
    PermisoCrearUsuario = permisos(4)
End Property
Public Property Get personal() As Boolean
    personal = permisos(5)
End Property
Public Function autenticacion(UserName As String, PASSWORD As String) As Boolean
    Dim DB_Usuario As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set DB_Usuario = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Dim consulta As String
    Dim ipRegistro As String
    Dim i As Integer
    
    For i = 0 To 5
        permisos(i) = False
    Next
    On Error GoTo fallo
    ipRegistro = ReadINI(App.Path + "\config.ini", "server", "ip")
    consulta = "SELECT * FROM empleados WHERE usuario='" & _
                UserName & "' AND password='" & _
                PASSWORD & "' AND anulado=0"
    DB_Usuario.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                                & "SERVER=" & ipRegistro & ";" _
                                & "DATABASE=geslab_usuarios;" _
                                & "UID=root;" _
                                & "PWD=;" _
                                & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    DB_Usuario.Open
    rs.ActiveConnection = DB_Usuario
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    rs.Open consulta
    If rs.RecordCount = 0 Then 'esto significa que NOrecuperado ningun registro
       MsgBox "La contraseña o el usuario no es válido. Vuelva a intentarlo.", vbExclamation, "Inicio de sesión"
       autenticacion = False
       estado1 = False
    Else
    'recoge todos los permisos de la tabla
    'permiso(0)=impresion // permiso(1)=facturacion //permiso(2)=modificacion
    'permiso(3)=eliminiacion) // permiso(4)=crear usuario // permiso(5)=personal
        If Trim(rs("uso")) <> "" Then
             MsgBox "El usuario ya esta en uso en la máquina " & rs("uso") & ".", vbExclamation, App.Title
'            autenticacion = False
'            estado1 = False
        End If
        For i = 0 To 5
           If rs(i + 5) = 1 Then
             permisos(i) = True
           End If
        Next
        datosUsuario(0) = rs(0) 'id
        datosUsuario(1) = rs(1) 'nombre
        datosUsuario(2) = rs(2) 'apellidos
        datosUsuario(3) = rs(3) 'usuario
'        consulta = "INSERT INTO login VALUES ('" & UserName & "','" & rs(2) & "','" & Now & "')"
'        DB_Usuario.Execute consulta
        Dim cadena As String
        cadena = frmLogin.Winsock1.LocalHostName
        If cadena = "" Then
         cadena = "No identificado"
        End If
        consulta = "UPDATE empleados set USO = '" & UCase(cadena) & "' where id_empleado=" & rs(0)
        DB_Usuario.Execute consulta
        autenticacion = True
        estado1 = True
    End If
    rs.Close
    DB_Usuario.Close
    Set rs = Nothing
    Set DB_Usuario = Nothing
    Exit Function
fallo:
    MsgBox Err.Number & "/" & Err.Description
    autenticacion = False
End Function

