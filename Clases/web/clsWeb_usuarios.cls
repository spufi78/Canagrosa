VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWeb_usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSWEB_USUARIOS
'***************************************************************
Private ID_USUARIO as Long
Private USUARIO as String
Private PASSWORD as String
Private NOMBRE as String
Private EMAIL as String
Private CLIENTE_ID as Long
Private TIPO_USUARIO_ID as Long
Private OFICINA_ID as Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSWEB_USUARIOS
'***************************************************************
Public Property Let setID_USUARIO(ByVal dato As Long)
	ID_USUARIO = dato 
End Property
Public Property Let setUSUARIO(ByVal dato As String)
	USUARIO = dato 
End Property
Public Property Let setPASSWORD(ByVal dato As String)
	PASSWORD = dato 
End Property
Public Property Let setNOMBRE(ByVal dato As String)
	NOMBRE = dato 
End Property
Public Property Let setEMAIL(ByVal dato As String)
	EMAIL = dato 
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
	CLIENTE_ID = dato 
End Property
Public Property Let setTIPO_USUARIO_ID(ByVal dato As Long)
	TIPO_USUARIO_ID = dato 
End Property
Public Property Let setOFICINA_ID(ByVal dato As Long)
	OFICINA_ID = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSWEB_USUARIOS
'***************************************************************
Public Property Get getID_USUARIO() As Long
	getID_USUARIO = ID_USUARIO
End Property
Public Property Get getUSUARIO() As String
	getUSUARIO = USUARIO
End Property
Public Property Get getPASSWORD() As String
	getPASSWORD = PASSWORD
End Property
Public Property Get getNOMBRE() As String
	getNOMBRE = NOMBRE
End Property
Public Property Get getEMAIL() As String
	getEMAIL = EMAIL
End Property
Public Property Get getCLIENTE_ID() As Long
	getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getTIPO_USUARIO_ID() As Long
	getTIPO_USUARIO_ID = TIPO_USUARIO_ID
End Property
Public Property Get getOFICINA_ID() As Long
	getOFICINA_ID = OFICINA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSWEB_USUARIOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM WEB_USUARIOS WHERE ID_USUARIO = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		ID_USUARIO = rs("ID_USUARIO")
		USUARIO = rs("USUARIO")
		PASSWORD = rs("PASSWORD")
		NOMBRE = rs("NOMBRE")
		EMAIL = rs("EMAIL")
		CLIENTE_ID = rs("CLIENTE_ID")
		TIPO_USUARIO_ID = rs("TIPO_USUARIO_ID")
		OFICINA_ID = rs("OFICINA_ID")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsWeb_usuarios)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
	On error goto fallo
	Dim consulta As String
	consulta = "INSERT INTO WEB_USUARIOS " & _
	           "  VALUES (" & _
			ID_USUARIO & "," & "'" & USUARIO & "'" & "," & "'" & PASSWORD & "'" & "," & _
			"'" & NOMBRE & "'" & "," & "'" & EMAIL & "'" & "," & CLIENTE_ID & "," & _
			TIPO_USUARIO_ID & "," & OFICINA_ID & _
		")"
	execute_bd consulta
	Insertar = ID_USUARIO
	Exit Function
fallo:
	Insertar = 0
	MsgBox "Error al insertar (clsWeb_usuarios)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "UPDATE WEB_USUARIOS SET " & _
			" USUARIO = '" & USUARIO & "'," & _
			" PASSWORD = '" & PASSWORD & "'," & _
			" NOMBRE = '" & NOMBRE & "'," & _
			" EMAIL = '" & EMAIL & "'," & _
			" CLIENTE_ID = " & CLIENTE_ID & "," & _
			" TIPO_USUARIO_ID = " & TIPO_USUARIO_ID & "," & _
			" OFICINA_ID = " & OFICINA_ID & "" & _
		" WHERE ID_USUARIO = " & ID
	execute_bd consulta
	Modificar = True
	Exit Function
fallo:
	Modificar = False
	MsgBox "Error al Modificar (clsWeb_usuarios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "DELETE FROM WEB_USUARIOS " & _
		"    WHERE ID_USUARIO = " & ID
	execute_bd consulta
	Eliminar = True
	Exit Function
fallo:
	Eliminar = False
	MsgBox "Error al Eliminar (clsWeb_usuarios)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM WEB_USUARIOS ORDER BY ID_USUARIO"
	Set Listado = datos_bd(consulta)
End Function
