VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPresupuestos_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPRESUPUESTOS_DETALLE
'***************************************************************
Private PRESUPUESTO_ID as Long
Private ORDEN as Long
Private DESCRIPCION as String
Private PRECIO as Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPRESUPUESTOS_DETALLE
'***************************************************************
Public Property Let setPRESUPUESTO_ID(ByVal dato As Long)
	PRESUPUESTO_ID = dato 
End Property
Public Property Let setORDEN(ByVal dato As Long)
	ORDEN = dato 
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
	DESCRIPCION = dato 
End Property
Public Property Let setPRECIO(ByVal dato As Long)
	PRECIO = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPRESUPUESTOS_DETALLE
'***************************************************************
Public Property Get getPRESUPUESTO_ID() As Long
	getPRESUPUESTO_ID = PRESUPUESTO_ID
End Property
Public Property Get getORDEN() As Long
	getORDEN = ORDEN
End Property
Public Property Get getDESCRIPCION() As String
	getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getPRECIO() As Long
	getPRECIO = PRECIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPRESUPUESTOS_DETALLE
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM PRESUPUESTOS_DETALLE WHERE PRESUPUESTO_ID = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		PRESUPUESTO_ID = rs("PRESUPUESTO_ID")
		ORDEN = rs("ORDEN")
		DESCRIPCION = rs("DESCRIPCION")
		PRECIO = rs("PRECIO")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsPresupuestos_detalle)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
	On error goto fallo
	Dim consulta As String
	consulta = "INSERT INTO PRESUPUESTOS_DETALLE " & _
	           "  VALUES (" & _
			PRESUPUESTO_ID & "," & ORDEN & "," & "'" & DESCRIPCION & "'" & "," & _
			PRECIO & _
		")"
	conn.Execute consulta
	Insertar = PRESUPUESTO_ID
	Exit Function
fallo:
	Insertar = 0
	MsgBox "Error al insertar (clsPresupuestos_detalle)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "UPDATE PRESUPUESTOS_DETALLE SET " & _
			" ORDEN = " & ORDEN & "," & _
			" DESCRIPCION = '" & DESCRIPCION & "'," & _
			" PRECIO = " & PRECIO & "" & _
		" WHERE PRESUPUESTO_ID = " & ID
	conn.Execute consulta
	Modificar = True
	Exit Function
fallo:
	Modificar = False
	MsgBox "Error al Modificar (clsPresupuestos_detalle)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "DELETE FROM PRESUPUESTOS_DETALLE " & _
		"    WHERE PRESUPUESTO_ID = " & ID
	conn.Execute consulta
	Eliminar = True
	Exit Function
fallo:
	Eliminar = False
	MsgBox "Error al Eliminar (clsPresupuestos_detalle)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM PRESUPUESTOS_DETALLE ORDER BY PRESUPUESTO_ID"
	Set Listado = datos_bd(consulta)
End Function
