VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensajes_usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMENSAJES_USUARIOS
'***************************************************************
Private MENSAJE_ID As Long
Private EMPLEADO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMENSAJES_USUARIOS
'***************************************************************
Public Property Let setMENSAJE_ID(ByVal dato As Long)
        MENSAJE_ID = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMENSAJES_USUARIOS
'***************************************************************
Public Property Get getMENSAJE_ID() As Long
        getMENSAJE_ID = MENSAJE_ID
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMENSAJES_USUARIOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM mensajes_usuarios WHERE MENSAJE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MENSAJE_ID = rs("MENSAJE_ID")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMensajes_usuarios)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO mensajes_usuarios " & _
                   "  VALUES (" & _
                        MENSAJE_ID & "," & EMPLEADO_ID & _
                ",0)"
        execute_bd consulta
        Insertar = MENSAJE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMensajes_usuarios)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE mensajes_usuarios SET " & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "" & _
                " WHERE MENSAJE_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMensajes_usuarios)", vbCritical, Err.Description
End Function
Public Function Leer(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE mensajes_usuarios SET " & _
                        " LEIDO = 1 " & _
                " WHERE MENSAJE_ID = " & ID & _
                "   AND EMPLEADO_ID = " & USUARIO.getID_EMPLEADO
        execute_bd consulta
        Leer = True
        Exit Function
fallo:
        Leer = False
        MsgBox "Error al Modificar (clsMensajes_usuarios)", vbCritical, Err.Description
End Function
Public Function Eliminar(mensaje As Long, USUARIO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM mensajes_usuarios " & _
                "    WHERE MENSAJE_ID = " & mensaje & _
                "      AND EMPLEADO_ID = " & USUARIO
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMensajes_usuarios)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.ID_EMPLEADO, B.NOMBRE, B.APELLIDOS, B.USUARIO " & _
                   "  FROM mensajes_usuarios A, usuarios B " & _
                   " WHERE A.EMPLEADO_ID = B.ID_EMPLEADO " & _
                   "   AND A.MENSAJE_ID = " & ID & _
                   " ORDER BY B.APELLIDOS, B.NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
