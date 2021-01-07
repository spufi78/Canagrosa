VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAu_programa_versiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSAU_PROGRAMA_VERSIONES
'***************************************************************
Private PROGRAMA_ID As Integer
Private VERSION As Integer
Private OBSERVACIONES As String
Private USUARIO_CREACION As Integer
Private FECHA_CREACION As String
Private USUARIO_APROBACION As Integer
Private FECHA_APROBACION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSAU_PROGRAMA_VERSIONES
'***************************************************************
Public Property Let setPROGRAMA_ID(ByVal dato As Integer)
        PROGRAMA_ID = dato
End Property
Public Property Let setVERSION(ByVal dato As Integer)
        VERSION = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setUSUARIO_CREACION(ByVal dato As Integer)
        USUARIO_CREACION = dato
End Property
Public Property Let setFECHA_CREACION(ByVal dato As String)
        FECHA_CREACION = dato
End Property
Public Property Let setUSUARIO_APROBACION(ByVal dato As Integer)
        USUARIO_APROBACION = dato
End Property
Public Property Let setFECHA_APROBACION(ByVal dato As String)
        FECHA_APROBACION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSAU_PROGRAMA_VERSIONES
'***************************************************************
Public Property Get getPROGRAMA_ID() As Integer
        getPROGRAMA_ID = PROGRAMA_ID
End Property
Public Property Get getVERSION() As Integer
        getVERSION = VERSION
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getUSUARIO_CREACION() As Integer
        getUSUARIO_CREACION = USUARIO_CREACION
End Property
Public Property Get getFECHA_CREACION() As String
        getFECHA_CREACION = FECHA_CREACION
End Property
Public Property Get getUSUARIO_APROBACION() As Integer
        getUSUARIO_APROBACION = USUARIO_APROBACION
End Property
Public Property Get getFECHA_APROBACION() As String
        getFECHA_APROBACION = FECHA_APROBACION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSAU_PROGRAMA_VERSIONES
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM AU_PROGRAMA_VERSIONES WHERE PROGRAMA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PROGRAMA_ID = rs("PROGRAMA_ID")
                VERSION = rs("VERSION")
                OBSERVACIONES = rs("OBSERVACIONES")
                USUARIO_CREACION = rs("USUARIO_CREACION")
                FECHA_CREACION = rs("FECHA_CREACION")
                USUARIO_APROBACION = rs("USUARIO_APROBACION")
                FECHA_APROBACION = rs("FECHA_APROBACION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAu_programa_versiones)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO AU_PROGRAMA_VERSIONES " & _
                   "  VALUES (" & _
                        PROGRAMA_ID & "," & VERSION & "," & "'" & OBSERVACIONES & "'" & "," & _
                        USUARIO_CREACION & "," & "'" & FECHA_CREACION & "'" & "," & USUARIO_APROBACION & "," & _
                        "'" & FECHA_APROBACION & "'" & _
                ")"
        execute_bd consulta
        Insertar = PROGRAMA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAu_programa_versiones)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Integer, VERSI As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE AU_PROGRAMA_VERSIONES SET " & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " USUARIO_CREACION = " & USUARIO_CREACION & "," & _
                        " FECHA_CREACION = '" & FECHA_CREACION & "'," & _
                        " USUARIO_APROBACION = " & USUARIO_APROBACION & "," & _
                        " FECHA_APROBACION = '" & FECHA_APROBACION & "'" & _
                   " WHERE PROGRAMA_ID = " & ID & _
                   "   AND VERSION = " & VERSI
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAu_programa_versiones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM AU_PROGRAMA_VERSIONES " & _
                "    WHERE PROGRAMA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAu_programa_versiones)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Integer) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT VERSION,OBSERVACIONES,USUARIO_CREACION,FECHA_CREACION,USUARIO_APROBACION,FECHA_APROBACION " & _
                   "  FROM AU_PROGRAMA_VERSIONES " & _
                   " WHERE PROGRAMA_ID = " & ID & _
                   " ORDER BY VERSION DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Aprobar(ID As Integer, VERSI As Integer, fecha As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE AU_PROGRAMA_VERSIONES SET " & _
                        " FECHA_APROBACION = '" & fecha & "'" & _
                   " WHERE PROGRAMA_ID = " & ID & _
                   "   AND VERSION = " & VERSI
        execute_bd consulta
        Aprobar = True
        Exit Function
fallo:
        Aprobar = False
        MsgBox "Error al Aprobar (clsAu_programa_versiones)", vbCritical, Err.Description
End Function


