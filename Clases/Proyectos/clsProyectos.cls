VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROYECTOS
'***************************************************************
Private ID_PROYECTO as Long
Private NOMBRE as String
Private FECHA_INICIO as String
Private FECHA_FIN as String
Private HORAS_PLANIFICADAS as Long
Private HORAS_DEDICADAS as Long
Private IMPORTE_PLANIFICADO as String
Private IMPOERTE_DEDICADO as String
Private COMENTARIO as String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROYECTOS
'***************************************************************
Public Property Let setID_PROYECTO(ByVal dato As Long)
	ID_PROYECTO = dato 
End Property
Public Property Let setNOMBRE(ByVal dato As String)
	NOMBRE = dato 
End Property
Public Property Let setFECHA_INICIO(ByVal dato As String)
	FECHA_INICIO = dato 
End Property
Public Property Let setFECHA_FIN(ByVal dato As String)
	FECHA_FIN = dato 
End Property
Public Property Let setHORAS_PLANIFICADAS(ByVal dato As Long)
	HORAS_PLANIFICADAS = dato 
End Property
Public Property Let setHORAS_DEDICADAS(ByVal dato As Long)
	HORAS_DEDICADAS = dato 
End Property
Public Property Let setIMPORTE_PLANIFICADO(ByVal dato As String)
	IMPORTE_PLANIFICADO = dato 
End Property
Public Property Let setIMPOERTE_DEDICADO(ByVal dato As String)
	IMPOERTE_DEDICADO = dato 
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
	COMENTARIO = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROYECTOS
'***************************************************************
Public Property Get getID_PROYECTO() As Long
	getID_PROYECTO = ID_PROYECTO
End Property
Public Property Get getNOMBRE() As String
	getNOMBRE = NOMBRE
End Property
Public Property Get getFECHA_INICIO() As String
	getFECHA_INICIO = FECHA_INICIO
End Property
Public Property Get getFECHA_FIN() As String
	getFECHA_FIN = FECHA_FIN
End Property
Public Property Get getHORAS_PLANIFICADAS() As Long
	getHORAS_PLANIFICADAS = HORAS_PLANIFICADAS
End Property
Public Property Get getHORAS_DEDICADAS() As Long
	getHORAS_DEDICADAS = HORAS_DEDICADAS
End Property
Public Property Get getIMPORTE_PLANIFICADO() As String
	getIMPORTE_PLANIFICADO = IMPORTE_PLANIFICADO
End Property
Public Property Get getIMPOERTE_DEDICADO() As String
	getIMPOERTE_DEDICADO = IMPOERTE_DEDICADO
End Property
Public Property Get getCOMENTARIO() As String
	getCOMENTARIO = COMENTARIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROYECTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM PROYECTOS WHERE ID_PROYECTO = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		ID_PROYECTO = rs("ID_PROYECTO")
		NOMBRE = rs("NOMBRE")
		FECHA_INICIO = rs("FECHA_INICIO")
		FECHA_FIN = rs("FECHA_FIN")
		HORAS_PLANIFICADAS = rs("HORAS_PLANIFICADAS")
		HORAS_DEDICADAS = rs("HORAS_DEDICADAS")
		IMPORTE_PLANIFICADO = rs("IMPORTE_PLANIFICADO")
		IMPOERTE_DEDICADO = rs("IMPOERTE_DEDICADO")
		COMENTARIO = rs("COMENTARIO")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsProyectos)", vbCritical, Err.Description
End Function
Public Function CrearID()
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT MAX(ID_PROYECTO) FROM PROYECTOS"
	Set rs = datos_bd(consulta)
	If Isnull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
		ID_PROYECTO = 1
	else
		ID_PROYECTO = rs.Fields(0) + 1
	End If
	Set rs = Nothing
End Function
Public Function Insertar() As Long
	On error goto fallo
	Dim consulta As String
	Me.CrearId
	consulta = "INSERT INTO PROYECTOS " & _
	           "  VALUES (" & _
			ID_PROYECTO & "," & "'" & NOMBRE & "'" & "," & "'" & FECHA_INICIO & "'" & "," & _
			"'" & FECHA_FIN & "'" & "," & HORAS_PLANIFICADAS & "," & HORAS_DEDICADAS & "," & _
			"'" & IMPORTE_PLANIFICADO & "'" & "," & "'" & IMPOERTE_DEDICADO & "'" & "," & "'" & COMENTARIO & "'" & _
		")"
	execute_bd consulta
	Insertar = ID_PROYECTO
	Exit Function
fallo:
	Insertar = 0
	MsgBox "Error al insertar (clsProyectos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "UPDATE PROYECTOS SET " & _
			" NOMBRE = '" & NOMBRE & "'," & _
			" FECHA_INICIO = '" & FECHA_INICIO & "'," & _
			" FECHA_FIN = '" & FECHA_FIN & "'," & _
			" HORAS_PLANIFICADAS = " & HORAS_PLANIFICADAS & "," & _
			" HORAS_DEDICADAS = " & HORAS_DEDICADAS & "," & _
			" IMPORTE_PLANIFICADO = '" & IMPORTE_PLANIFICADO & "'," & _
			" IMPOERTE_DEDICADO = '" & IMPOERTE_DEDICADO & "'," & _
			" COMENTARIO = '" & COMENTARIO & "'" & _
		" WHERE ID_PROYECTO = " & ID
	execute_bd consulta
	Modificar = True
	Exit Function
fallo:
	Modificar = False
	MsgBox "Error al Modificar (clsProyectos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "DELETE FROM PROYECTOS " & _
		"    WHERE ID_PROYECTO = " & ID
	execute_bd consulta
	Eliminar = True
	Exit Function
fallo:
	Eliminar = False
	MsgBox "Error al Eliminar (clsProyectos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM PROYECTOS ORDER BY ID_PROYECTO"
	Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT ID_PROYECTO,NOMBRE FROM PROYECTOS ORDER BY NOMBRE"
	Set Listado_Combo = datos_bd(consulta)
End Function
