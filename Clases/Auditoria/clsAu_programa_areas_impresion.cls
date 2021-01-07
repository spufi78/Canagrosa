VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAu_programa_areas_impresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSAU_PROGRAMA_AREAS_IMPRESION
'***************************************************************
Private PROGRAMA_ID as Long
Private AREA_ID as Long
Private M1 as Long
Private M2 as Long
Private M3 as Long
Private M4 as Long
Private M5 as Long
Private M6 as Long
Private M7 as Long
Private M8 as Long
Private M9 as Long
Private M10 as Long
Private M11 as Long
Private M12 as Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSAU_PROGRAMA_AREAS_IMPRESION
'***************************************************************
Public Property Let setPROGRAMA_ID(ByVal dato As Long)
	PROGRAMA_ID = dato 
End Property
Public Property Let setAREA_ID(ByVal dato As Long)
	AREA_ID = dato 
End Property
Public Property Let setM1(ByVal dato As Long)
	M1 = dato 
End Property
Public Property Let setM2(ByVal dato As Long)
	M2 = dato 
End Property
Public Property Let setM3(ByVal dato As Long)
	M3 = dato 
End Property
Public Property Let setM4(ByVal dato As Long)
	M4 = dato 
End Property
Public Property Let setM5(ByVal dato As Long)
	M5 = dato 
End Property
Public Property Let setM6(ByVal dato As Long)
	M6 = dato 
End Property
Public Property Let setM7(ByVal dato As Long)
	M7 = dato 
End Property
Public Property Let setM8(ByVal dato As Long)
	M8 = dato 
End Property
Public Property Let setM9(ByVal dato As Long)
	M9 = dato 
End Property
Public Property Let setM10(ByVal dato As Long)
	M10 = dato 
End Property
Public Property Let setM11(ByVal dato As Long)
	M11 = dato 
End Property
Public Property Let setM12(ByVal dato As Long)
	M12 = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSAU_PROGRAMA_AREAS_IMPRESION
'***************************************************************
Public Property Get getPROGRAMA_ID() As Long
	getPROGRAMA_ID = PROGRAMA_ID
End Property
Public Property Get getAREA_ID() As Long
	getAREA_ID = AREA_ID
End Property
Public Property Get getM1() As Long
	getM1 = M1
End Property
Public Property Get getM2() As Long
	getM2 = M2
End Property
Public Property Get getM3() As Long
	getM3 = M3
End Property
Public Property Get getM4() As Long
	getM4 = M4
End Property
Public Property Get getM5() As Long
	getM5 = M5
End Property
Public Property Get getM6() As Long
	getM6 = M6
End Property
Public Property Get getM7() As Long
	getM7 = M7
End Property
Public Property Get getM8() As Long
	getM8 = M8
End Property
Public Property Get getM9() As Long
	getM9 = M9
End Property
Public Property Get getM10() As Long
	getM10 = M10
End Property
Public Property Get getM11() As Long
	getM11 = M11
End Property
Public Property Get getM12() As Long
	getM12 = M12
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSAU_PROGRAMA_AREAS_IMPRESION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM AU_PROGRAMA_AREAS_IMPRESION WHERE PROGRAMA_ID = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		PROGRAMA_ID = rs("PROGRAMA_ID")
		AREA_ID = rs("AREA_ID")
		M1 = rs("M1")
		M2 = rs("M2")
		M3 = rs("M3")
		M4 = rs("M4")
		M5 = rs("M5")
		M6 = rs("M6")
		M7 = rs("M7")
		M8 = rs("M8")
		M9 = rs("M9")
		M10 = rs("M10")
		M11 = rs("M11")
		M12 = rs("M12")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsAu_programa_areas_impresion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
	On error goto fallo
	Dim consulta As String
	consulta = "INSERT INTO AU_PROGRAMA_AREAS_IMPRESION " & _
	           "  VALUES (" & _
			PROGRAMA_ID & "," & AREA_ID & "," & M1 & "," & _
			M2 & "," & M3 & "," & M4 & "," & _
			M5 & "," & M6 & "," & M7 & "," & _
			M8 & "," & M9 & "," & M10 & "," & _
			M11 & "," & M12 & _
		")"
	execute_bd consulta
	Insertar = PROGRAMA_ID
	Exit Function
fallo:
	Insertar = 0
	MsgBox "Error al insertar (clsAu_programa_areas_impresion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "UPDATE AU_PROGRAMA_AREAS_IMPRESION SET " & _
			" AREA_ID = " & AREA_ID & "," & _
			" M1 = " & M1 & "," & _
			" M2 = " & M2 & "," & _
			" M3 = " & M3 & "," & _
			" M4 = " & M4 & "," & _
			" M5 = " & M5 & "," & _
			" M6 = " & M6 & "," & _
			" M7 = " & M7 & "," & _
			" M8 = " & M8 & "," & _
			" M9 = " & M9 & "," & _
			" M10 = " & M10 & "," & _
			" M11 = " & M11 & "," & _
			" M12 = " & M12 & "" & _
		" WHERE PROGRAMA_ID = " & ID
	execute_bd consulta
	Modificar = True
	Exit Function
fallo:
	Modificar = False
	MsgBox "Error al Modificar (clsAu_programa_areas_impresion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
	On error goto fallo
	Dim consulta As String
	consulta = "DELETE FROM AU_PROGRAMA_AREAS_IMPRESION " & _
		"    WHERE PROGRAMA_ID = " & ID
	execute_bd consulta
	Eliminar = True
	Exit Function
fallo:
	Eliminar = False
	MsgBox "Error al Eliminar (clsAu_programa_areas_impresion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM AU_PROGRAMA_AREAS_IMPRESION ORDER BY PROGRAMA_ID"
	Set Listado = datos_bd(consulta)
End Function
