VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAu_programa_areas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSAU_PROGRAMA_AREAS
'***************************************************************
Private PROGRAMA_ID As Integer
Private AREA_ID As Integer
Private FECHA_PREVISTA As String
Private FECHA_REALIZACION As String
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSAU_PROGRAMA_AREAS
'***************************************************************
Public Property Let setPROGRAMA_ID(ByVal dato As Integer)
        PROGRAMA_ID = dato
End Property
Public Property Let setAREA_ID(ByVal dato As Integer)
        AREA_ID = dato
End Property
Public Property Let setFECHA_PREVISTA(ByVal dato As String)
        FECHA_PREVISTA = dato
End Property
Public Property Let setFECHA_REALIZACION(ByVal dato As String)
        FECHA_REALIZACION = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSAU_PROGRAMA_AREAS
'***************************************************************
Public Property Get getPROGRAMA_ID() As Integer
        getPROGRAMA_ID = PROGRAMA_ID
End Property
Public Property Get getAREA_ID() As Integer
        getAREA_ID = AREA_ID
End Property
Public Property Get getFECHA_PREVISTA() As String
        getFECHA_PREVISTA = FECHA_PREVISTA
End Property
Public Property Get getFECHA_REALIZACION() As String
        getFECHA_REALIZACION = FECHA_REALIZACION
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSAU_PROGRAMA_AREAS
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM AU_PROGRAMA_AREAS WHERE PROGRAMA_ID = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                PROGRAMA_ID = RS("PROGRAMA_ID")
                AREA_ID = RS("AREA_ID")
                FECHA_PREVISTA = RS("FECHA_PREVISTA")
                FECHA_REALIZACION = RS("FECHA_REALIZACION")
                ORDEN = RS("ORDEN")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAu_programa_areas)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO AU_PROGRAMA_AREAS " & _
                   "  VALUES (" & _
                        PROGRAMA_ID & "," & AREA_ID & "," & "'" & FECHA_PREVISTA & "'," & FECHA_REALIZACION & "," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = PROGRAMA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAu_programa_areas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE AU_PROGRAMA_AREAS SET " & _
                        " AREA_ID = " & AREA_ID & "," & _
                        " FECHA_PREVISTA = '" & FECHA_PREVISTA & "'," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE PROGRAMA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAu_programa_areas)", vbCritical, Err.Description
End Function
Public Function ModificarRealizacion(PROGRAMA As Integer, AREA As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE AU_PROGRAMA_AREAS " & _
                   "   SET FECHA_REALIZACION = " & FECHA_REALIZACION & _
                   " WHERE PROGRAMA_ID = " & PROGRAMA & _
                   "   AND AREA_ID = " & AREA
        execute_bd consulta
        ModificarRealizacion = True
        Exit Function
fallo:
        ModificarRealizacion = False
        MsgBox "Error al ModificarRealizacion (clsAu_programa_areas)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM AU_PROGRAMA_AREAS " & _
                "    WHERE PROGRAMA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAu_programa_areas)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.AREA_ID,B.AREA,A.FECHA_PREVISTA,A.FECHA_REALIZACION " & _
                   "  FROM AU_PROGRAMA_AREAS A, AU_AREAS B " & _
                   " WHERE A.PROGRAMA_ID = " & ID & _
                   "   AND A.AREA_ID = B.ID_AREA " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
