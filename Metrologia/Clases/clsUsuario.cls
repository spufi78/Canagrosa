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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private ID_EMPLEADO As Long
Private NOMBRE As String
Private USUARIO As String
Private PASSWORD As String
Private PER_1 As Integer
Private PER_2 As Integer
Private PER_3 As Integer
Private PER_4 As Integer
Private PER_5 As Integer
Private PER_6 As Integer
Private PER_7 As Integer
Public Property Let setID_EMPLEADO(ByVal dato As Long)
    ID_EMPLEADO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setUSUARIO(ByVal dato As String)
    USUARIO = Trim(dato)
End Property
Public Property Let setPASSWORD(ByVal dato As String)
    PASSWORD = Trim(dato)
End Property
Public Property Let setPER_1(ByVal dato As Integer)
    PER_1 = dato
End Property
Public Property Let setPER_2(ByVal dato As Integer)
    PER_2 = dato
End Property
Public Property Let setPER_3(ByVal dato As Integer)
    PER_3 = dato
End Property
Public Property Let setPER_4(ByVal dato As Integer)
    PER_4 = dato
End Property
Public Property Let setPER_5(ByVal dato As Integer)
    PER_5 = dato
End Property
Public Property Let setPER_6(ByVal dato As Integer)
    PER_6 = dato
End Property
Public Property Let setPER_7(ByVal dato As Integer)
    PER_7 = dato
End Property
Public Property Get getID_EMPLEADO() As Long
    getID_EMPLEADO = ID_EMPLEADO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getUSUARIO() As String
    getUSUARIO = USUARIO
End Property
Public Property Get getPASSWORD() As String
    getPASSWORD = PASSWORD
End Property
Public Property Get getPER_1() As Integer
    getPER_1 = PER_1
End Property
Public Property Get getPER_2() As Integer
    getPER_2 = PER_2
End Property
Public Property Get getPER_3() As Integer
    getPER_3 = PER_3
End Property
Public Property Get getPER_4() As Integer
    getPER_4 = PER_4
End Property
Public Property Get getPER_5() As Integer
    getPER_5 = PER_5
End Property
Public Property Get getPER_6() As Integer
    getPER_6 = PER_6
End Property
Public Property Get getPER_7() As Integer
    getPER_7 = PER_7
End Property
Public Function Cargar(ByVal cod_usuario As Long) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM USUARIOS WHERE ID_EMPLEADO =" & cod_usuario
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        ID_EMPLEADO = rs("ID_EMPLEADO")
        NOMBRE = rs("NOMBRE")
        USUARIO = rs("USUARIO")
        PASSWORD = rs("password")
        PER_1 = rs("PER_1")
        PER_2 = rs("PER_2")
        PER_3 = rs("PER_3")
        PER_4 = rs("PER_4")
        PER_5 = rs("PER_5")
        PER_6 = rs("PER_6")
        PER_7 = rs("PER_7")
        Cargar = True
    Else
        Cargar = False
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
fallo:
    Cargar = False
    MsgBox "Error al recuperar el usuario_", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM USUARIOS order by id_empleado"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_EMPLEADO, NOMBRE FROM USUARIOS " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function

Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "DELETE FROM USUARIOS WHERE id_empleado = " & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    MsgBox Err.Number & "/" & Err.Description
    Eliminar = False
End Function
Public Sub CrearID()
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_EMPLEADO) FROM USUARIOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_EMPLEADO = 1
    Else
        ID_EMPLEADO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox Err.Number & "/" & Err.Description
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    If ID_EMPLEADO = 0 Then
    Me.CrearID
    End If
    Dim consulta As String
    consulta = "Insert into USUARIOS " & _
               " values(" & _
               ID_EMPLEADO & ",'" & NOMBRE & "','" & _
               USUARIO & "','" & PASSWORD & "'," & _
               PER_1 & "," & _
               PER_2 & "," & _
               PER_3 & "," & _
               PER_4 & "," & _
               PER_5 & "," & _
               PER_6 & "," & _
               PER_7 & ")"
    execute_bd consulta
    Insertar = ID_EMPLEADO
    Exit Function
fallo:
    MsgBox "Error al insertar el usuario (Insertar)", vbCritical, Err.Description
    Insertar = 0
End Function
Public Function Modificar(USU As Long) As Long
    On Error GoTo fallo
    Dim consulta As String
    Dim ousu As New ClsUsuario
    With ousu
    .setID_EMPLEADO = USU
    .setUSUARIO = USUARIO
    .setNOMBRE = NOMBRE
    .setPASSWORD = PASSWORD
    .setPER_1 = PER_1
    .setPER_2 = PER_2
    .setPER_3 = PER_3
    .setPER_4 = PER_4
    .setPER_5 = PER_5
    .setPER_6 = PER_6
    .setPER_7 = PER_7
    .Eliminar (USU)
    .Insertar
    End With
    Modificar = USU
    Exit Function
fallo:
    MsgBox "Error al modificar el usuario (Modificar)", vbCritical, Err.Description
    Modificar = 0
End Function
Public Function autentificacion(UserName As String, PASSWORD As String) As Integer
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    On Error GoTo fallo
    consulta = "SELECT ID_EMPLEADO FROM USUARIOS " & _
               " WHERE usuario='" & UserName & "' " & _
               "   AND password='" & PASSWORD & "'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
       autentificacion = 0
    Else
       autentificacion = rs(0)
    End If
    Set rs = Nothing
    Exit Function
fallo:
    MsgBox Err.Number & "/" & Err.Description
    autentificacion = 0
End Function
