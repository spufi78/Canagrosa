VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentos_liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDOCUMENTOS_LIQUIDACION
'***************************************************************
Private DOCUMENTO_ID as Long
Private COMERCIAL_ID as Long
Private LIQUIDADO as Long
Private FECHA_LIQUIDACION as String
Private COMISION as Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDOCUMENTOS_LIQUIDACION
'***************************************************************
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
	DOCUMENTO_ID = dato 
End Property
Public Property Let setCOMERCIAL_ID(ByVal dato As Long)
	COMERCIAL_ID = dato 
End Property
Public Property Let setLIQUIDADO(ByVal dato As Long)
	LIQUIDADO = dato 
End Property
Public Property Let setFECHA_LIQUIDACION(ByVal dato As String)
	FECHA_LIQUIDACION = dato 
End Property
Public Property Let setCOMISION(ByVal dato As Long)
	COMISION = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDOCUMENTOS_LIQUIDACION
'***************************************************************
Public Property Get getDOCUMENTO_ID() As Long
	getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getCOMERCIAL_ID() As Long
	getCOMERCIAL_ID = COMERCIAL_ID
End Property
Public Property Get getLIQUIDADO() As Long
	getLIQUIDADO = LIQUIDADO
End Property
Public Property Get getFECHA_LIQUIDACION() As String
	getFECHA_LIQUIDACION = FECHA_LIQUIDACION
End Property
Public Property Get getCOMISION() As Long
	getCOMISION = COMISION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDOCUMENTOS_LIQUIDACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM DOCUMENTOS_LIQUIDACION WHERE DOCUMENTO_ID = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		DOCUMENTO_ID = rs("DOCUMENTO_ID")
		COMERCIAL_ID = rs("COMERCIAL_ID")
		LIQUIDADO = rs("LIQUIDADO")
		FECHA_LIQUIDACION = rs("FECHA_LIQUIDACION")
		COMISION = rs("COMISION")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsDocumentos_liquidacion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
	On error goto fallo
	Dim consulta As String
	consulta = "INSERT INTO DOCUMENTOS_LIQUIDACION " & _
	           "  VALUES (" & _
			DOCUMENTO_ID & "," & COMERCIAL_ID & "," & LIQUIDADO & "," & _
			"'" & FECHA_LIQUIDACION & "'" & "," & COMISION & _
		")"
	execute_bd consulta
	Insertar = DOCUMENTO_ID
	Exit Function
fallo:
	Insertar = 0
	MsgBox "Error al insertar (clsDocumentos_liquidacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "UPDATE DOCUMENTOS_LIQUIDACION SET " & _
			" COMERCIAL_ID = " & COMERCIAL_ID & "," & _
			" LIQUIDADO = " & LIQUIDADO & "," & _
			" FECHA_LIQUIDACION = '" & FECHA_LIQUIDACION & "'," & _
			" COMISION = " & COMISION & "" & _
		" WHERE DOCUMENTO_ID = " & ID
	execute_bd consulta
	Modificar = True
	Exit Function
fallo:
	Modificar = False
	MsgBox "Error al Modificar (clsDocumentos_liquidacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "DELETE FROM DOCUMENTOS_LIQUIDACION " & _
		"    WHERE DOCUMENTO_ID = " & ID
	execute_bd consulta
	Eliminar = True
	Exit Function
fallo:
	Eliminar = False
	MsgBox "Error al Eliminar (clsDocumentos_liquidacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM DOCUMENTOS_LIQUIDACION ORDER BY DOCUMENTO_ID"
	Set Listado = datos_bd(consulta)
End Function
