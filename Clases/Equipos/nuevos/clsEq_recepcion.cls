VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEq_recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsEq_recepcion
'***************************************************************
Private EQUIPO_ID as Long
Private USUARIO_ID as Long
Private FECHA as String
Private OP1 as Long
Private OP2 as Long
Private OP3 as Long
Private OP4 as Long
Private OP5 as Long
Private OP6 as Long
Private OP7 as String
Private OP8 as Long
Private OP9 as Long
Private OP10 as String
Private OBSERVACIONES as String
Private OPCONFORME as Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsEq_recepcion
'***************************************************************
Public Property Let setEQUIPO_ID(ByVal dato As Long)
	EQUIPO_ID = dato 
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
	USUARIO_ID = dato 
End Property
Public Property Let setFECHA(ByVal dato As String)
	FECHA = dato 
End Property
Public Property Let setOP1(ByVal dato As Long)
	OP1 = dato 
End Property
Public Property Let setOP2(ByVal dato As Long)
	OP2 = dato 
End Property
Public Property Let setOP3(ByVal dato As Long)
	OP3 = dato 
End Property
Public Property Let setOP4(ByVal dato As Long)
	OP4 = dato 
End Property
Public Property Let setOP5(ByVal dato As Long)
	OP5 = dato 
End Property
Public Property Let setOP6(ByVal dato As Long)
	OP6 = dato 
End Property
Public Property Let setOP7(ByVal dato As String)
	OP7 = dato 
End Property
Public Property Let setOP8(ByVal dato As Long)
	OP8 = dato 
End Property
Public Property Let setOP9(ByVal dato As Long)
	OP9 = dato 
End Property
Public Property Let setOP10(ByVal dato As String)
	OP10 = dato 
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
	OBSERVACIONES = dato 
End Property
Public Property Let setOPCONFORME(ByVal dato As Long)
	OPCONFORME = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsEq_recepcion
'***************************************************************
Public Property Get getEQUIPO_ID() As Long
	getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getUSUARIO_ID() As Long
	getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFECHA() As String
	getFECHA = FECHA
End Property
Public Property Get getOP1() As Long
	getOP1 = OP1
End Property
Public Property Get getOP2() As Long
	getOP2 = OP2
End Property
Public Property Get getOP3() As Long
	getOP3 = OP3
End Property
Public Property Get getOP4() As Long
	getOP4 = OP4
End Property
Public Property Get getOP5() As Long
	getOP5 = OP5
End Property
Public Property Get getOP6() As Long
	getOP6 = OP6
End Property
Public Property Get getOP7() As String
	getOP7 = OP7
End Property
Public Property Get getOP8() As Long
	getOP8 = OP8
End Property
Public Property Get getOP9() As Long
	getOP9 = OP9
End Property
Public Property Get getOP10() As String
	getOP10 = OP10
End Property
Public Property Get getOBSERVACIONES() As String
	getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getOPCONFORME() As Long
	getOPCONFORME = OPCONFORME
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsEq_recepcion
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM EQ_RECEPCION WHERE EQUIPO_ID = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		EQUIPO_ID = rs("EQUIPO_ID")
		USUARIO_ID = rs("USUARIO_ID")
		FECHA = rs("FECHA")
		OP1 = rs("OP1")
		OP2 = rs("OP2")
		OP3 = rs("OP3")
		OP4 = rs("OP4")
		OP5 = rs("OP5")
		OP6 = rs("OP6")
		OP7 = rs("OP7")
		OP8 = rs("OP8")
		OP9 = rs("OP9")
		OP10 = rs("OP10")
		OBSERVACIONES = rs("OBSERVACIONES")
		OPCONFORME = rs("OPCONFORME")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsEq_recepcion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
	On error goto fallo
	Dim consulta As String
	consulta = "INSERT INTO EQ_RECEPCION " & _
	           "  VALUES (" & _
			EQUIPO_ID & "," & USUARIO_ID & "," & "'" & FECHA & "'" & "," & _
			OP1 & "," & OP2 & "," & OP3 & "," & _
			OP4 & "," & OP5 & "," & OP6 & "," & _
			"'" & OP7 & "'" & "," & OP8 & "," & OP9 & "," & _
			"'" & OP10 & "'" & "," & "'" & OBSERVACIONES & "'" & "," & OPCONFORME & _
		")"
	execute_bd consulta
	Insertar = EQUIPO_ID
	Exit Function
fallo:
	Insertar = 0
	MsgBox "Error al insertar (clsEq_recepcion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "UPDATE EQ_RECEPCION SET " & _
			" USUARIO_ID = " & USUARIO_ID & "," & _
			" FECHA = '" & FECHA & "'," & _
			" OP1 = " & OP1 & "," & _
			" OP2 = " & OP2 & "," & _
			" OP3 = " & OP3 & "," & _
			" OP4 = " & OP4 & "," & _
			" OP5 = " & OP5 & "," & _
			" OP6 = " & OP6 & "," & _
			" OP7 = '" & OP7 & "'," & _
			" OP8 = " & OP8 & "," & _
			" OP9 = " & OP9 & "," & _
			" OP10 = '" & OP10 & "'," & _
			" OBSERVACIONES = '" & OBSERVACIONES & "'," & _
			" OPCONFORME = " & OPCONFORME & "" & _
		" WHERE EQUIPO_ID = " & ID
	execute_bd consulta
	Modificar = True
	Exit Function
fallo:
	Modificar = False
	MsgBox "Error al Modificar (clsEq_recepcion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "DELETE FROM EQ_RECEPCION " & _
		"    WHERE EQUIPO_ID = " & ID
	execute_bd consulta
	Eliminar = True
	Exit Function
fallo:
	Eliminar = False
	MsgBox "Error al Eliminar (clsEq_recepcion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM EQ_RECEPCION ORDER BY EQUIPO_ID"
	Set Listado = datos_bd(consulta)
End Function
