VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_historico_curso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_HISTORICO_CURSO
'***************************************************************
Private CURSO_ID as Long
Private MOVIMIENTO as Long
Private DESCRIPCION as String
Private FECHA_PARADA as String
Private USUARIO as Long
Private FTIMESTP as Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_HISTORICO_CURSO
'***************************************************************
Public Property Let setCURSO_ID(ByVal dato As Long)
	CURSO_ID = dato 
End Property
Public Property Let setMOVIMIENTO(ByVal dato As Long)
	MOVIMIENTO = dato 
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
	DESCRIPCION = dato 
End Property
Public Property Let setFECHA_PARADA(ByVal dato As String)
	FECHA_PARADA = dato 
End Property
Public Property Let setUSUARIO(ByVal dato As Long)
	USUARIO = dato 
End Property
Public Property Let setFTIMESTP(ByVal dato As Long)
	FTIMESTP = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_HISTORICO_CURSO
'***************************************************************
Public Property Get getCURSO_ID() As Long
	getCURSO_ID = CURSO_ID
End Property
Public Property Get getMOVIMIENTO() As Long
	getMOVIMIENTO = MOVIMIENTO
End Property
Public Property Get getDESCRIPCION() As String
	getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA_PARADA() As String
	getFECHA_PARADA = FECHA_PARADA
End Property
Public Property Get getUSUARIO() As Long
	getUSUARIO = USUARIO
End Property
Public Property Get getFTIMESTP() As Long
	getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_HISTORICO_CURSO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM FORMACION_HISTORICO_CURSO WHERE CURSO_ID = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		CURSO_ID = rs("CURSO_ID")
		MOVIMIENTO = rs("MOVIMIENTO")
		DESCRIPCION = rs("DESCRIPCION")
		FECHA_PARADA = rs("FECHA_PARADA")
		USUARIO = rs("USUARIO")
		FTIMESTP = rs("FTIMESTP")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsFormacion_historico_curso)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
	On error goto fallo
	Dim consulta As String
	consulta = "INSERT INTO FORMACION_HISTORICO_CURSO " & _
	           "  VALUES (" & _
			CURSO_ID & "," & MOVIMIENTO & "," & "'" & DESCRIPCION & "'" & "," & _
			"'" & FECHA_PARADA & "'" & "," & USUARIO & "," & FTIMESTP & _
		")"
	execute_bd consulta
	Insertar = CURSO_ID
	Exit Function
fallo:
	Insertar = 0
	MsgBox "Error al insertar (clsFormacion_historico_curso)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "UPDATE FORMACION_HISTORICO_CURSO SET " & _
			" MOVIMIENTO = " & MOVIMIENTO & "," & _
			" DESCRIPCION = '" & DESCRIPCION & "'," & _
			" FECHA_PARADA = '" & FECHA_PARADA & "'," & _
			" USUARIO = " & USUARIO & "," & _
			" FTIMESTP = " & FTIMESTP & "" & _
		" WHERE CURSO_ID = " & ID
	execute_bd consulta
	Modificar = True
	Exit Function
fallo:
	Modificar = False
	MsgBox "Error al Modificar (clsFormacion_historico_curso)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "DELETE FROM FORMACION_HISTORICO_CURSO " & _
		"    WHERE CURSO_ID = " & ID
	execute_bd consulta
	Eliminar = True
	Exit Function
fallo:
	Eliminar = False
	MsgBox "Error al Eliminar (clsFormacion_historico_curso)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM FORMACION_HISTORICO_CURSO ORDER BY CURSO_ID"
	Set Listado = datos_bd(consulta)
End Function
