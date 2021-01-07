VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContenedor_modelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCONTENEDOR_MODELOS
'***************************************************************
Private ID_MODELO_CONTENEDOR as Long
Private NOMBRE as String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCONTENEDOR_MODELOS
'***************************************************************
Public Property Let setID_MODELO_CONTENEDOR(ByVal dato As Long)
	ID_MODELO_CONTENEDOR = dato 
End Property
Public Property Let setNOMBRE(ByVal dato As String)
	NOMBRE = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCONTENEDOR_MODELOS
'***************************************************************
Public Property Get getID_MODELO_CONTENEDOR() As Long
	getID_MODELO_CONTENEDOR = ID_MODELO_CONTENEDOR
End Property
Public Property Get getNOMBRE() As String
	getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCONTENEDOR_MODELOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM CONTENEDOR_MODELOS WHERE ID_MODELO_CONTENEDOR = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		ID_MODELO_CONTENEDOR = rs("ID_MODELO_CONTENEDOR")
		NOMBRE = rs("NOMBRE")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsContenedor_modelos)", vbCritical, Err.Description
End Function
Public Function CrearID()
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT MAX(ID_MODELO_CONTENEDOR) FROM CONTENEDOR_MODELOS"
	Set rs = datos_bd(consulta)
	If Isnull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
		ID_MODELO_CONTENEDOR = 1
	else
		ID_MODELO_CONTENEDOR = rs.Fields(0) + 1
	End If
	Set rs = Nothing
End Function
Public Function Insertar() As Long
	On error goto fallo
	Dim consulta As String
	Me.CrearId
	consulta = "INSERT INTO CONTENEDOR_MODELOS " & _
	           "  VALUES (" & _
			ID_MODELO_CONTENEDOR & "," & "'" & NOMBRE & "'" & _
		")"
	execute_bd consulta
	Insertar = ID_MODELO_CONTENEDOR
	Exit Function
fallo:
	Insertar = 0
	MsgBox "Error al insertar (clsContenedor_modelos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "UPDATE CONTENEDOR_MODELOS SET " & _
			" NOMBRE = '" & NOMBRE & "'" & _
		" WHERE ID_MODELO_CONTENEDOR = " & ID
	execute_bd consulta
	Modificar = True
	Exit Function
fallo:
	Modificar = False
	MsgBox "Error al Modificar (clsContenedor_modelos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "DELETE FROM CONTENEDOR_MODELOS " & _
		"    WHERE ID_MODELO_CONTENEDOR = " & ID
	execute_bd consulta
	Eliminar = True
	Exit Function
fallo:
	Eliminar = False
	MsgBox "Error al Eliminar (clsContenedor_modelos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM CONTENEDOR_MODELOS ORDER BY ID_MODELO_CONTENEDOR"
	Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT ID_MODELO_CONTENEDOR,NOMBRE FROM CONTENEDOR_MODELOS ORDER BY NOMBRE"
	Set Listado_Combo = datos_bd(consulta)
End Function
