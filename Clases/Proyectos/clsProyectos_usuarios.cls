VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProyectos_usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROYECTOS_USUARIOS
'***************************************************************
Private PROYECTO_ID as Long
Private EMPLEADO_ID as Long
Private HORAS as Long
Private ANULADO as Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROYECTOS_USUARIOS
'***************************************************************
Public Property Let setPROYECTO_ID(ByVal dato As Long)
	PROYECTO_ID = dato 
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
	EMPLEADO_ID = dato 
End Property
Public Property Let setHORAS(ByVal dato As Long)
	HORAS = dato 
End Property
Public Property Let setANULADO(ByVal dato As Long)
	ANULADO = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROYECTOS_USUARIOS
'***************************************************************
Public Property Get getPROYECTO_ID() As Long
	getPROYECTO_ID = PROYECTO_ID
End Property
Public Property Get getEMPLEADO_ID() As Long
	getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getHORAS() As Long
	getHORAS = HORAS
End Property
Public Property Get getANULADO() As Long
	getANULADO = ANULADO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROYECTOS_USUARIOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM PROYECTOS_USUARIOS WHERE PROYECTO_ID = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		PROYECTO_ID = rs("PROYECTO_ID")
		EMPLEADO_ID = rs("EMPLEADO_ID")
		HORAS = rs("HORAS")
		ANULADO = rs("ANULADO")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsProyectos_usuarios)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
	On error goto fallo
	Dim consulta As String
	consulta = "INSERT INTO PROYECTOS_USUARIOS " & _
	           "  VALUES (" & _
			PROYECTO_ID & "," & EMPLEADO_ID & "," & HORAS & "," & _
			ANULADO & _
		")"
	execute_bd consulta
	Insertar = PROYECTO_ID
	Exit Function
fallo:
	Insertar = 0
	MsgBox "Error al insertar (clsProyectos_usuarios)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "UPDATE PROYECTOS_USUARIOS SET " & _
			" EMPLEADO_ID = " & EMPLEADO_ID & "," & _
			" HORAS = " & HORAS & "," & _
			" ANULADO = " & ANULADO & "" & _
		" WHERE PROYECTO_ID = " & ID
	execute_bd consulta
	Modificar = True
	Exit Function
fallo:
	Modificar = False
	MsgBox "Error al Modificar (clsProyectos_usuarios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "DELETE FROM PROYECTOS_USUARIOS " & _
		"    WHERE PROYECTO_ID = " & ID
	execute_bd consulta
	Eliminar = True
	Exit Function
fallo:
	Eliminar = False
	MsgBox "Error al Eliminar (clsProyectos_usuarios)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM PROYECTOS_USUARIOS ORDER BY PROYECTO_ID"
	Set Listado = datos_bd(consulta)
End Function
