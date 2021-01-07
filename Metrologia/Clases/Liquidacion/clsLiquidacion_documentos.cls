VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLiquidacion_documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSLIQUIDACION_DOCUMENTOS
'***************************************************************
Private LIQUIDACION_ID As Long
Private DOCUMENTO_ID As Long
Private COMISION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSLIQUIDACION_DOCUMENTOS
'***************************************************************
Public Property Let setLIQUIDACION_ID(ByVal dato As Long)
        LIQUIDACION_ID = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setCOMISION(ByVal dato As String)
        COMISION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSLIQUIDACION_DOCUMENTOS
'***************************************************************
Public Property Get getLIQUIDACION_ID() As Long
        getLIQUIDACION_ID = LIQUIDACION_ID
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getCOMISION() As String
        getCOMISION = COMISION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSLIQUIDACION_DOCUMENTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM LIQUIDACION_DOCUMENTOS WHERE LIQUIDACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                LIQUIDACION_ID = rs("LIQUIDACION_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                COMISION = rs("COMISION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsLiquidacion_documentos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO LIQUIDACION_DOCUMENTOS " & _
                   "  VALUES (" & _
                        LIQUIDACION_ID & "," & DOCUMENTO_ID & ",'" & COMISION & "'" & _
                ")"
        execute_bd consulta
        Insertar = LIQUIDACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsLiquidacion_documentos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE LIQUIDACION_DOCUMENTOS SET " & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " COMISION = '" & COMISION & "'" & _
                " WHERE LIQUIDACION_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsLiquidacion_documentos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM LIQUIDACION_DOCUMENTOS " & _
                "    WHERE LIQUIDACION_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsLiquidacion_documentos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM LIQUIDACION_DOCUMENTOS ORDER BY LIQUIDACION_ID"
        Set Listado = datos_bd(consulta)
End Function
