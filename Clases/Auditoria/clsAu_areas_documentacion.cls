VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAu_areas_documentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSAU_AREAS_DOCUMENTACION
'***************************************************************
Private AREA_ID As Integer
Private DOCUMENTO_ID As Integer
Private tipo As Integer
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSAU_AREAS_DOCUMENTACION
'***************************************************************
Public Property Let setAREA_ID(ByVal dato As Integer)
        AREA_ID = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Integer)
        DOCUMENTO_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As Integer)
        tipo = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSAU_AREAS_DOCUMENTACION
'***************************************************************
Public Property Get getAREA_ID() As Integer
        getAREA_ID = AREA_ID
End Property
Public Property Get getDOCUMENTO_ID() As Integer
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getTIPO() As Integer
        getTIPO = tipo
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSAU_AREAS_DOCUMENTACION
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM AU_AREAS_DOCUMENTACION WHERE AREA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                AREA_ID = rs("AREA_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                tipo = rs("TIPO")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAu_areas_documentacion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO AU_AREAS_DOCUMENTACION " & _
                   "  VALUES (" & _
                        AREA_ID & "," & DOCUMENTO_ID & "," & tipo & "," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = AREA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAu_areas_documentacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE AU_AREAS_DOCUMENTACION SET " & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " TIPO = " & tipo & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE AREA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAu_areas_documentacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM AU_AREAS_DOCUMENTACION " & _
                "    WHERE AREA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAu_areas_documentacion)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Integer) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM AU_AREAS_DOCUMENTACION " & _
                   " WHERE AREA_ID = " & ID & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function
