VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLecturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSLECTURAS
'***************************************************************
Private CodigoAbonado as Long
Private Anyo as Long
Private L1 as String
Private L2 as String
Private L3 as String
Private L4 as String
Private L5 as String
Private L6 as String
Private L7 as String
Private L8 as String
Private L9 as String
Private L10 as String
Private L11 as String
Private L12 as String
Private C1 as String
Private C2 as String
Private C3 as String
Private C4 as String
Private C5 as String
Private C6 as String
Private C7 as String
Private C8 as String
Private C9 as String
Private C10 as String
Private C11 as String
Private C12 as String
Private F1 as String
Private F2 as String
Private F3 as String
Private F4 as String
Private F5 as String
Private F6 as String
Private F7 as String
Private F8 as String
Private F9 as String
Private F10 as String
Private F11 as String
Private F12 as String
Private LA1 as String
Private LA2 as String
Private LA3 as String
Private LA4 as String
Private LA5 as String
Private LA6 as String
Private LA7 as String
Private LA8 as String
Private LA9 as String
Private LA10 as String
Private LA11 as String
Private LA12 as String
Private FechaL1 as String
Private FechaL2 as String
Private FechaL3 as String
Private FechaL4 as String
Private FechaL5 as String
Private FechaL6 as String
Private FechaL7 as String
Private FechaL8 as String
Private FechaL9 as String
Private FechaL10 as String
Private FechaL11 as String
Private FechaL12 as String
Private FechaLA1 as String
Private FechaLA2 as String
Private FechaLA3 as String
Private FechaLA4 as String
Private FechaLA5 as String
Private FechaLA6 as String
Private FechaLA7 as String
Private FechaLA8 as String
Private FechaLA9 as String
Private FechaLA10 as String
Private FechaLA11 as String
Private FechaLA12 as String
Private Inc1 as String
Private Inc2 as String
Private Inc3 as String
Private Inc4 as String
Private Inc5 as String
Private Inc6 as String
Private Inc7 as String
Private Inc8 as String
Private Inc9 as String
Private Inc10 as String
Private Inc11 as String
Private Inc12 as String
Private CI1 as String
Private CI2 as String
Private CI3 as String
Private CI4 as String
Private CI5 as String
Private CI6 as String
Private CI7 as String
Private CI8 as String
Private CI9 as String
Private CI10 as String
Private CI11 as String
Private CI12 as String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSLECTURAS
'***************************************************************
Public Property Let setCODIGOABONADO(ByVal dato As Long)
	CODIGOABONADO = dato 
End Property
Public Property Let setANYO(ByVal dato As Long)
	ANYO = dato 
End Property
Public Property Let setL1(ByVal dato As String)
	L1 = dato 
End Property
Public Property Let setL2(ByVal dato As String)
	L2 = dato 
End Property
Public Property Let setL3(ByVal dato As String)
	L3 = dato 
End Property
Public Property Let setL4(ByVal dato As String)
	L4 = dato 
End Property
Public Property Let setL5(ByVal dato As String)
	L5 = dato 
End Property
Public Property Let setL6(ByVal dato As String)
	L6 = dato 
End Property
Public Property Let setL7(ByVal dato As String)
	L7 = dato 
End Property
Public Property Let setL8(ByVal dato As String)
	L8 = dato 
End Property
Public Property Let setL9(ByVal dato As String)
	L9 = dato 
End Property
Public Property Let setL10(ByVal dato As String)
	L10 = dato 
End Property
Public Property Let setL11(ByVal dato As String)
	L11 = dato 
End Property
Public Property Let setL12(ByVal dato As String)
	L12 = dato 
End Property
Public Property Let setC1(ByVal dato As String)
	C1 = dato 
End Property
Public Property Let setC2(ByVal dato As String)
	C2 = dato 
End Property
Public Property Let setC3(ByVal dato As String)
	C3 = dato 
End Property
Public Property Let setC4(ByVal dato As String)
	C4 = dato 
End Property
Public Property Let setC5(ByVal dato As String)
	C5 = dato 
End Property
Public Property Let setC6(ByVal dato As String)
	C6 = dato 
End Property
Public Property Let setC7(ByVal dato As String)
	C7 = dato 
End Property
Public Property Let setC8(ByVal dato As String)
	C8 = dato 
End Property
Public Property Let setC9(ByVal dato As String)
	C9 = dato 
End Property
Public Property Let setC10(ByVal dato As String)
	C10 = dato 
End Property
Public Property Let setC11(ByVal dato As String)
	C11 = dato 
End Property
Public Property Let setC12(ByVal dato As String)
	C12 = dato 
End Property
Public Property Let setF1(ByVal dato As String)
	F1 = dato 
End Property
Public Property Let setF2(ByVal dato As String)
	F2 = dato 
End Property
Public Property Let setF3(ByVal dato As String)
	F3 = dato 
End Property
Public Property Let setF4(ByVal dato As String)
	F4 = dato 
End Property
Public Property Let setF5(ByVal dato As String)
	F5 = dato 
End Property
Public Property Let setF6(ByVal dato As String)
	F6 = dato 
End Property
Public Property Let setF7(ByVal dato As String)
	F7 = dato 
End Property
Public Property Let setF8(ByVal dato As String)
	F8 = dato 
End Property
Public Property Let setF9(ByVal dato As String)
	F9 = dato 
End Property
Public Property Let setF10(ByVal dato As String)
	F10 = dato 
End Property
Public Property Let setF11(ByVal dato As String)
	F11 = dato 
End Property
Public Property Let setF12(ByVal dato As String)
	F12 = dato 
End Property
Public Property Let setLA1(ByVal dato As String)
	LA1 = dato 
End Property
Public Property Let setLA2(ByVal dato As String)
	LA2 = dato 
End Property
Public Property Let setLA3(ByVal dato As String)
	LA3 = dato 
End Property
Public Property Let setLA4(ByVal dato As String)
	LA4 = dato 
End Property
Public Property Let setLA5(ByVal dato As String)
	LA5 = dato 
End Property
Public Property Let setLA6(ByVal dato As String)
	LA6 = dato 
End Property
Public Property Let setLA7(ByVal dato As String)
	LA7 = dato 
End Property
Public Property Let setLA8(ByVal dato As String)
	LA8 = dato 
End Property
Public Property Let setLA9(ByVal dato As String)
	LA9 = dato 
End Property
Public Property Let setLA10(ByVal dato As String)
	LA10 = dato 
End Property
Public Property Let setLA11(ByVal dato As String)
	LA11 = dato 
End Property
Public Property Let setLA12(ByVal dato As String)
	LA12 = dato 
End Property
Public Property Let setFECHAL1(ByVal dato As String)
	FECHAL1 = dato 
End Property
Public Property Let setFECHAL2(ByVal dato As String)
	FECHAL2 = dato 
End Property
Public Property Let setFECHAL3(ByVal dato As String)
	FECHAL3 = dato 
End Property
Public Property Let setFECHAL4(ByVal dato As String)
	FECHAL4 = dato 
End Property
Public Property Let setFECHAL5(ByVal dato As String)
	FECHAL5 = dato 
End Property
Public Property Let setFECHAL6(ByVal dato As String)
	FECHAL6 = dato 
End Property
Public Property Let setFECHAL7(ByVal dato As String)
	FECHAL7 = dato 
End Property
Public Property Let setFECHAL8(ByVal dato As String)
	FECHAL8 = dato 
End Property
Public Property Let setFECHAL9(ByVal dato As String)
	FECHAL9 = dato 
End Property
Public Property Let setFECHAL10(ByVal dato As String)
	FECHAL10 = dato 
End Property
Public Property Let setFECHAL11(ByVal dato As String)
	FECHAL11 = dato 
End Property
Public Property Let setFECHAL12(ByVal dato As String)
	FECHAL12 = dato 
End Property
Public Property Let setFECHALA1(ByVal dato As String)
	FECHALA1 = dato 
End Property
Public Property Let setFECHALA2(ByVal dato As String)
	FECHALA2 = dato 
End Property
Public Property Let setFECHALA3(ByVal dato As String)
	FECHALA3 = dato 
End Property
Public Property Let setFECHALA4(ByVal dato As String)
	FECHALA4 = dato 
End Property
Public Property Let setFECHALA5(ByVal dato As String)
	FECHALA5 = dato 
End Property
Public Property Let setFECHALA6(ByVal dato As String)
	FECHALA6 = dato 
End Property
Public Property Let setFECHALA7(ByVal dato As String)
	FECHALA7 = dato 
End Property
Public Property Let setFECHALA8(ByVal dato As String)
	FECHALA8 = dato 
End Property
Public Property Let setFECHALA9(ByVal dato As String)
	FECHALA9 = dato 
End Property
Public Property Let setFECHALA10(ByVal dato As String)
	FECHALA10 = dato 
End Property
Public Property Let setFECHALA11(ByVal dato As String)
	FECHALA11 = dato 
End Property
Public Property Let setFECHALA12(ByVal dato As String)
	FECHALA12 = dato 
End Property
Public Property Let setINC1(ByVal dato As String)
	INC1 = dato 
End Property
Public Property Let setINC2(ByVal dato As String)
	INC2 = dato 
End Property
Public Property Let setINC3(ByVal dato As String)
	INC3 = dato 
End Property
Public Property Let setINC4(ByVal dato As String)
	INC4 = dato 
End Property
Public Property Let setINC5(ByVal dato As String)
	INC5 = dato 
End Property
Public Property Let setINC6(ByVal dato As String)
	INC6 = dato 
End Property
Public Property Let setINC7(ByVal dato As String)
	INC7 = dato 
End Property
Public Property Let setINC8(ByVal dato As String)
	INC8 = dato 
End Property
Public Property Let setINC9(ByVal dato As String)
	INC9 = dato 
End Property
Public Property Let setINC10(ByVal dato As String)
	INC10 = dato 
End Property
Public Property Let setINC11(ByVal dato As String)
	INC11 = dato 
End Property
Public Property Let setINC12(ByVal dato As String)
	INC12 = dato 
End Property
Public Property Let setCI1(ByVal dato As String)
	CI1 = dato 
End Property
Public Property Let setCI2(ByVal dato As String)
	CI2 = dato 
End Property
Public Property Let setCI3(ByVal dato As String)
	CI3 = dato 
End Property
Public Property Let setCI4(ByVal dato As String)
	CI4 = dato 
End Property
Public Property Let setCI5(ByVal dato As String)
	CI5 = dato 
End Property
Public Property Let setCI6(ByVal dato As String)
	CI6 = dato 
End Property
Public Property Let setCI7(ByVal dato As String)
	CI7 = dato 
End Property
Public Property Let setCI8(ByVal dato As String)
	CI8 = dato 
End Property
Public Property Let setCI9(ByVal dato As String)
	CI9 = dato 
End Property
Public Property Let setCI10(ByVal dato As String)
	CI10 = dato 
End Property
Public Property Let setCI11(ByVal dato As String)
	CI11 = dato 
End Property
Public Property Let setCI12(ByVal dato As String)
	CI12 = dato 
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSLECTURAS
'***************************************************************
Public Property Get getCODIGOABONADO() As Long
	getCODIGOABONADO = CodigoAbonado
End Property
Public Property Get getANYO() As Long
	getANYO = Anyo
End Property
Public Property Get getL1() As String
	getL1 = L1
End Property
Public Property Get getL2() As String
	getL2 = L2
End Property
Public Property Get getL3() As String
	getL3 = L3
End Property
Public Property Get getL4() As String
	getL4 = L4
End Property
Public Property Get getL5() As String
	getL5 = L5
End Property
Public Property Get getL6() As String
	getL6 = L6
End Property
Public Property Get getL7() As String
	getL7 = L7
End Property
Public Property Get getL8() As String
	getL8 = L8
End Property
Public Property Get getL9() As String
	getL9 = L9
End Property
Public Property Get getL10() As String
	getL10 = L10
End Property
Public Property Get getL11() As String
	getL11 = L11
End Property
Public Property Get getL12() As String
	getL12 = L12
End Property
Public Property Get getC1() As String
	getC1 = C1
End Property
Public Property Get getC2() As String
	getC2 = C2
End Property
Public Property Get getC3() As String
	getC3 = C3
End Property
Public Property Get getC4() As String
	getC4 = C4
End Property
Public Property Get getC5() As String
	getC5 = C5
End Property
Public Property Get getC6() As String
	getC6 = C6
End Property
Public Property Get getC7() As String
	getC7 = C7
End Property
Public Property Get getC8() As String
	getC8 = C8
End Property
Public Property Get getC9() As String
	getC9 = C9
End Property
Public Property Get getC10() As String
	getC10 = C10
End Property
Public Property Get getC11() As String
	getC11 = C11
End Property
Public Property Get getC12() As String
	getC12 = C12
End Property
Public Property Get getF1() As String
	getF1 = F1
End Property
Public Property Get getF2() As String
	getF2 = F2
End Property
Public Property Get getF3() As String
	getF3 = F3
End Property
Public Property Get getF4() As String
	getF4 = F4
End Property
Public Property Get getF5() As String
	getF5 = F5
End Property
Public Property Get getF6() As String
	getF6 = F6
End Property
Public Property Get getF7() As String
	getF7 = F7
End Property
Public Property Get getF8() As String
	getF8 = F8
End Property
Public Property Get getF9() As String
	getF9 = F9
End Property
Public Property Get getF10() As String
	getF10 = F10
End Property
Public Property Get getF11() As String
	getF11 = F11
End Property
Public Property Get getF12() As String
	getF12 = F12
End Property
Public Property Get getLA1() As String
	getLA1 = LA1
End Property
Public Property Get getLA2() As String
	getLA2 = LA2
End Property
Public Property Get getLA3() As String
	getLA3 = LA3
End Property
Public Property Get getLA4() As String
	getLA4 = LA4
End Property
Public Property Get getLA5() As String
	getLA5 = LA5
End Property
Public Property Get getLA6() As String
	getLA6 = LA6
End Property
Public Property Get getLA7() As String
	getLA7 = LA7
End Property
Public Property Get getLA8() As String
	getLA8 = LA8
End Property
Public Property Get getLA9() As String
	getLA9 = LA9
End Property
Public Property Get getLA10() As String
	getLA10 = LA10
End Property
Public Property Get getLA11() As String
	getLA11 = LA11
End Property
Public Property Get getLA12() As String
	getLA12 = LA12
End Property
Public Property Get getFECHAL1() As String
	getFECHAL1 = FechaL1
End Property
Public Property Get getFECHAL2() As String
	getFECHAL2 = FechaL2
End Property
Public Property Get getFECHAL3() As String
	getFECHAL3 = FechaL3
End Property
Public Property Get getFECHAL4() As String
	getFECHAL4 = FechaL4
End Property
Public Property Get getFECHAL5() As String
	getFECHAL5 = FechaL5
End Property
Public Property Get getFECHAL6() As String
	getFECHAL6 = FechaL6
End Property
Public Property Get getFECHAL7() As String
	getFECHAL7 = FechaL7
End Property
Public Property Get getFECHAL8() As String
	getFECHAL8 = FechaL8
End Property
Public Property Get getFECHAL9() As String
	getFECHAL9 = FechaL9
End Property
Public Property Get getFECHAL10() As String
	getFECHAL10 = FechaL10
End Property
Public Property Get getFECHAL11() As String
	getFECHAL11 = FechaL11
End Property
Public Property Get getFECHAL12() As String
	getFECHAL12 = FechaL12
End Property
Public Property Get getFECHALA1() As String
	getFECHALA1 = FechaLA1
End Property
Public Property Get getFECHALA2() As String
	getFECHALA2 = FechaLA2
End Property
Public Property Get getFECHALA3() As String
	getFECHALA3 = FechaLA3
End Property
Public Property Get getFECHALA4() As String
	getFECHALA4 = FechaLA4
End Property
Public Property Get getFECHALA5() As String
	getFECHALA5 = FechaLA5
End Property
Public Property Get getFECHALA6() As String
	getFECHALA6 = FechaLA6
End Property
Public Property Get getFECHALA7() As String
	getFECHALA7 = FechaLA7
End Property
Public Property Get getFECHALA8() As String
	getFECHALA8 = FechaLA8
End Property
Public Property Get getFECHALA9() As String
	getFECHALA9 = FechaLA9
End Property
Public Property Get getFECHALA10() As String
	getFECHALA10 = FechaLA10
End Property
Public Property Get getFECHALA11() As String
	getFECHALA11 = FechaLA11
End Property
Public Property Get getFECHALA12() As String
	getFECHALA12 = FechaLA12
End Property
Public Property Get getINC1() As String
	getINC1 = Inc1
End Property
Public Property Get getINC2() As String
	getINC2 = Inc2
End Property
Public Property Get getINC3() As String
	getINC3 = Inc3
End Property
Public Property Get getINC4() As String
	getINC4 = Inc4
End Property
Public Property Get getINC5() As String
	getINC5 = Inc5
End Property
Public Property Get getINC6() As String
	getINC6 = Inc6
End Property
Public Property Get getINC7() As String
	getINC7 = Inc7
End Property
Public Property Get getINC8() As String
	getINC8 = Inc8
End Property
Public Property Get getINC9() As String
	getINC9 = Inc9
End Property
Public Property Get getINC10() As String
	getINC10 = Inc10
End Property
Public Property Get getINC11() As String
	getINC11 = Inc11
End Property
Public Property Get getINC12() As String
	getINC12 = Inc12
End Property
Public Property Get getCI1() As String
	getCI1 = CI1
End Property
Public Property Get getCI2() As String
	getCI2 = CI2
End Property
Public Property Get getCI3() As String
	getCI3 = CI3
End Property
Public Property Get getCI4() As String
	getCI4 = CI4
End Property
Public Property Get getCI5() As String
	getCI5 = CI5
End Property
Public Property Get getCI6() As String
	getCI6 = CI6
End Property
Public Property Get getCI7() As String
	getCI7 = CI7
End Property
Public Property Get getCI8() As String
	getCI8 = CI8
End Property
Public Property Get getCI9() As String
	getCI9 = CI9
End Property
Public Property Get getCI10() As String
	getCI10 = CI10
End Property
Public Property Get getCI11() As String
	getCI11 = CI11
End Property
Public Property Get getCI12() As String
	getCI12 = CI12
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSLECTURAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
	On error goto fallo
	Dim rs As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM LECTURAS WHERE CodigoAbonado = " & ID
	Set rs = datos_bd(consulta)
	If rs.RecordCount = 0 Then
		Carga = False
	else
		CODIGOABONADO = rs("CodigoAbonado")
		ANYO = rs("Anyo")
		L1 = rs("L1")
		L2 = rs("L2")
		L3 = rs("L3")
		L4 = rs("L4")
		L5 = rs("L5")
		L6 = rs("L6")
		L7 = rs("L7")
		L8 = rs("L8")
		L9 = rs("L9")
		L10 = rs("L10")
		L11 = rs("L11")
		L12 = rs("L12")
		C1 = rs("C1")
		C2 = rs("C2")
		C3 = rs("C3")
		C4 = rs("C4")
		C5 = rs("C5")
		C6 = rs("C6")
		C7 = rs("C7")
		C8 = rs("C8")
		C9 = rs("C9")
		C10 = rs("C10")
		C11 = rs("C11")
		C12 = rs("C12")
		F1 = rs("F1")
		F2 = rs("F2")
		F3 = rs("F3")
		F4 = rs("F4")
		F5 = rs("F5")
		F6 = rs("F6")
		F7 = rs("F7")
		F8 = rs("F8")
		F9 = rs("F9")
		F10 = rs("F10")
		F11 = rs("F11")
		F12 = rs("F12")
		LA1 = rs("LA1")
		LA2 = rs("LA2")
		LA3 = rs("LA3")
		LA4 = rs("LA4")
		LA5 = rs("LA5")
		LA6 = rs("LA6")
		LA7 = rs("LA7")
		LA8 = rs("LA8")
		LA9 = rs("LA9")
		LA10 = rs("LA10")
		LA11 = rs("LA11")
		LA12 = rs("LA12")
		FECHAL1 = rs("FechaL1")
		FECHAL2 = rs("FechaL2")
		FECHAL3 = rs("FechaL3")
		FECHAL4 = rs("FechaL4")
		FECHAL5 = rs("FechaL5")
		FECHAL6 = rs("FechaL6")
		FECHAL7 = rs("FechaL7")
		FECHAL8 = rs("FechaL8")
		FECHAL9 = rs("FechaL9")
		FECHAL10 = rs("FechaL10")
		FECHAL11 = rs("FechaL11")
		FECHAL12 = rs("FechaL12")
		FECHALA1 = rs("FechaLA1")
		FECHALA2 = rs("FechaLA2")
		FECHALA3 = rs("FechaLA3")
		FECHALA4 = rs("FechaLA4")
		FECHALA5 = rs("FechaLA5")
		FECHALA6 = rs("FechaLA6")
		FECHALA7 = rs("FechaLA7")
		FECHALA8 = rs("FechaLA8")
		FECHALA9 = rs("FechaLA9")
		FECHALA10 = rs("FechaLA10")
		FECHALA11 = rs("FechaLA11")
		FECHALA12 = rs("FechaLA12")
		INC1 = rs("Inc1")
		INC2 = rs("Inc2")
		INC3 = rs("Inc3")
		INC4 = rs("Inc4")
		INC5 = rs("Inc5")
		INC6 = rs("Inc6")
		INC7 = rs("Inc7")
		INC8 = rs("Inc8")
		INC9 = rs("Inc9")
		INC10 = rs("Inc10")
		INC11 = rs("Inc11")
		INC12 = rs("Inc12")
		CI1 = rs("CI1")
		CI2 = rs("CI2")
		CI3 = rs("CI3")
		CI4 = rs("CI4")
		CI5 = rs("CI5")
		CI6 = rs("CI6")
		CI7 = rs("CI7")
		CI8 = rs("CI8")
		CI9 = rs("CI9")
		CI10 = rs("CI10")
		CI11 = rs("CI11")
		CI12 = rs("CI12")
		Carga = True
	End if
	Set rs = Nothing
	Exit Function
fallo:
	Carga = False
	MsgBox "Error al cargar los datos(clsLecturas)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
	Dim consulta As String
	consulta = "SELECT * FROM LECTURAS ORDER BY CodigoAbonado"
	Set Listado = datos_bd(consulta)
End Function
