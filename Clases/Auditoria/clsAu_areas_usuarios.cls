VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAu_areas_usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSAU_AREAS_USUARIOS
'***************************************************************
Private AREA_ID As Integer
Private USUARIO_ID As Integer
Private TIPO_USUARIO As Integer
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSAU_AREAS_USUARIOS
'***************************************************************
Public Property Let setAREA_ID(ByVal dato As Integer)
        AREA_ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Integer)
        USUARIO_ID = dato
End Property
Public Property Let setTIPO_USUARIO(ByVal dato As Integer)
        TIPO_USUARIO = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSAU_AREAS_USUARIOS
'***************************************************************
Public Property Get getAREA_ID() As Integer
        getAREA_ID = AREA_ID
End Property
Public Property Get getUSUARIO_ID() As Integer
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTIPO_USUARIO() As Integer
        getTIPO_USUARIO = TIPO_USUARIO
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSAU_AREAS_USUARIOS
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM AU_AREAS_USUARIOS WHERE AREA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                AREA_ID = rs("AREA_ID")
                USUARIO_ID = rs("USUARIO_ID")
                TIPO_USUARIO = rs("TIPO_USUARIO")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAu_areas_usuarios)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO AU_AREAS_USUARIOS " & _
                   "  VALUES (" & _
                        AREA_ID & "," & USUARIO_ID & "," & TIPO_USUARIO & "," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = AREA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAu_areas_usuarios)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE AU_AREAS_USUARIOS SET " & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TIPO_USUARIO = " & TIPO_USUARIO & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE AREA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAu_areas_usuarios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM AU_AREAS_USUARIOS " & _
                "    WHERE AREA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAu_areas_usuarios)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.USUARIO_ID,concat(B.NOMBRE,' ',B.APELLIDOS), A.TIPO_USUARIO, B.EMAIL " & _
                   "  FROM AU_AREAS_USUARIOS A,usuarios B " & _
                   " WHERE A.USUARIO_ID = B.ID_EMPLEADO " & _
                   "   AND A.AREA_ID = " & ID & _
                   " ORDER BY TIPO_USUARIO,ORDEN"
        Set Listado = datos_bd(consulta)
End Function
