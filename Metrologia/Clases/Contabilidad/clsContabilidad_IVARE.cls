VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad_IVARE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsContabilidad_IVARE
'***************************************************************
Private Nº_Apunte As String
Private Fecha_contable As String
Private Nº_factura As String
Private Registro As String
Private Cuenta As String
Private Titulo_c As String
Private Nif As String
Private Fecha_factura As String
Private Basei1 As String
Private Porcei1 As String
Private Cuotai1 As String
Private Porcer1 As String
Private Cuotar1 As String
Private Basei2 As String
Private Porcei2 As String
Private Cuotai2 As String
Private Porcer2 As String
Private Cuotar2 As String
Private Basei3 As String
Private Porcei3 As String
Private Cuotai3 As String
Private Porcer3 As String
Private Cuotar3 As String
Private Basei4 As String
Private Porcei4 As String
Private Cuotai4 As String
Private Porcer4 As String
Private Cuotar4 As String
Private Base_exenta As String
Private Total_factura As String
Private OPERADOR As String
Private CP As String
Private cr As String
Private co As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsContabilidad_IVARE
'***************************************************************
Public Property Let setNº_APUNTE(ByVal dato As String)
        Nº_Apunte = dato
End Property
Public Property Let setFECHA_CONTABLE(ByVal dato As String)
        Fecha_contable = dato
End Property
Public Property Let setNº_FACTURA(ByVal dato As String)
        Nº_factura = dato
End Property
Public Property Let setREGISTRO(ByVal dato As String)
        Registro = dato
End Property
Public Property Let setCUENTA(ByVal dato As String)
        Cuenta = dato
End Property
Public Property Let setTITULO_C(ByVal dato As String)
        Titulo_c = dato
End Property
Public Property Let setNIF(ByVal dato As String)
        Nif = dato
End Property
Public Property Let setFECHA_FACTURA(ByVal dato As String)
        Fecha_factura = dato
End Property
Public Property Let setBASEI1(ByVal dato As String)
        Basei1 = dato
End Property
Public Property Let setPORCEI1(ByVal dato As String)
        Porcei1 = dato
End Property
Public Property Let setCUOTAI1(ByVal dato As String)
        Cuotai1 = dato
End Property
Public Property Let setPORCER1(ByVal dato As String)
        Porcer1 = dato
End Property
Public Property Let setCUOTAR1(ByVal dato As String)
        Cuotar1 = dato
End Property
Public Property Let setBASEI2(ByVal dato As String)
        Basei2 = dato
End Property
Public Property Let setPORCEI2(ByVal dato As String)
        Porcei2 = dato
End Property
Public Property Let setCUOTAI2(ByVal dato As String)
        Cuotai2 = dato
End Property
Public Property Let setPORCER2(ByVal dato As String)
        Porcer2 = dato
End Property
Public Property Let setCUOTAR2(ByVal dato As String)
        Cuotar2 = dato
End Property
Public Property Let setBASEI3(ByVal dato As String)
        Basei3 = dato
End Property
Public Property Let setPORCEI3(ByVal dato As String)
        Porcei3 = dato
End Property
Public Property Let setCUOTAI3(ByVal dato As String)
        Cuotai3 = dato
End Property
Public Property Let setPORCER3(ByVal dato As String)
        Porcer3 = dato
End Property
Public Property Let setCUOTAR3(ByVal dato As String)
        Cuotar3 = dato
End Property
Public Property Let setBASEI4(ByVal dato As String)
        Basei4 = dato
End Property
Public Property Let setPORCEI4(ByVal dato As String)
        Porcei4 = dato
End Property
Public Property Let setCUOTAI4(ByVal dato As String)
        Cuotai4 = dato
End Property
Public Property Let setPORCER4(ByVal dato As String)
        Porcer4 = dato
End Property
Public Property Let setCUOTAR4(ByVal dato As String)
        Cuotar4 = dato
End Property
Public Property Let setBASE_EXENTA(ByVal dato As String)
        Base_exenta = dato
End Property
Public Property Let setTOTAL_FACTURA(ByVal dato As String)
        Total_factura = dato
End Property
Public Property Let setOPERADOR(ByVal dato As String)
        OPERADOR = dato
End Property
Public Property Let setCP(ByVal dato As String)
        CP = dato
End Property
Public Property Let setCR(ByVal dato As String)
        cr = dato
End Property
Public Property Let setCO(ByVal dato As String)
        co = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsContabilidad_IVARE
'***************************************************************
Public Property Get getNº_APUNTE() As String
        getNº_APUNTE = Nº_Apunte
End Property
Public Property Get getFECHA_CONTABLE() As String
        getFECHA_CONTABLE = Fecha_contable
End Property
Public Property Get getNº_FACTURA() As String
        getNº_FACTURA = Nº_factura
End Property
Public Property Get getREGISTRO() As String
        getREGISTRO = Registro
End Property
Public Property Get getCUENTA() As String
        getCUENTA = Cuenta
End Property
Public Property Get getTITULO_C() As String
        getTITULO_C = Titulo_c
End Property
Public Property Get getNIF() As String
        getNIF = Nif
End Property
Public Property Get getFECHA_FACTURA() As String
        getFECHA_FACTURA = Fecha_factura
End Property
Public Property Get getBASEI1() As String
        getBASEI1 = Basei1
End Property
Public Property Get getPORCEI1() As String
        getPORCEI1 = Porcei1
End Property
Public Property Get getCUOTAI1() As String
        getCUOTAI1 = Cuotai1
End Property
Public Property Get getPORCER1() As String
        getPORCER1 = Porcer1
End Property
Public Property Get getCUOTAR1() As String
        getCUOTAR1 = Cuotar1
End Property
Public Property Get getBASEI2() As String
        getBASEI2 = Basei2
End Property
Public Property Get getPORCEI2() As String
        getPORCEI2 = Porcei2
End Property
Public Property Get getCUOTAI2() As String
        getCUOTAI2 = Cuotai2
End Property
Public Property Get getPORCER2() As String
        getPORCER2 = Porcer2
End Property
Public Property Get getCUOTAR2() As String
        getCUOTAR2 = Cuotar2
End Property
Public Property Get getBASEI3() As String
        getBASEI3 = Basei3
End Property
Public Property Get getPORCEI3() As String
        getPORCEI3 = Porcei3
End Property
Public Property Get getCUOTAI3() As String
        getCUOTAI3 = Cuotai3
End Property
Public Property Get getPORCER3() As String
        getPORCER3 = Porcer3
End Property
Public Property Get getCUOTAR3() As String
        getCUOTAR3 = Cuotar3
End Property
Public Property Get getBASEI4() As String
        getBASEI4 = Basei4
End Property
Public Property Get getPORCEI4() As String
        getPORCEI4 = Porcei4
End Property
Public Property Get getCUOTAI4() As String
        getCUOTAI4 = Cuotai4
End Property
Public Property Get getPORCER4() As String
        getPORCER4 = Porcer4
End Property
Public Property Get getCUOTAR4() As String
        getCUOTAR4 = Cuotar4
End Property
Public Property Get getBASE_EXENTA() As String
        getBASE_EXENTA = Base_exenta
End Property
Public Property Get getTOTAL_FACTURA() As String
        getTOTAL_FACTURA = Total_factura
End Property
Public Property Get getOPERADOR() As String
        getOPERADOR = OPERADOR
End Property
Public Property Get getCP() As String
        getCP = CP
End Property
Public Property Get getCR() As String
        getCR = cr
End Property
Public Property Get getCO() As String
        getCO = co
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsContabilidad_IVARE
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM IVARE WHERE Nº_Apunte = " & ID
        Dim oC As New clsContabilidad
        Set rs = oC.datos_bd_contabilidad(consulta)
        Set oC = Nothing
        
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                Nº_Apunte = rs("Nº_Apunte")
                Fecha_contable = rs("Fecha_contable")
                Nº_factura = rs("Nº_factura")
                Registro = rs("Registro")
                Cuenta = rs("Cuenta")
                Titulo_c = rs("Titulo_c")
                Nif = rs("Nif")
                Fecha_factura = rs("Fecha_factura")
                Basei1 = rs("Basei1")
                Porcei1 = rs("Porcei1")
                Cuotai1 = rs("Cuotai1")
                Porcer1 = rs("Porcer1")
                Cuotar1 = rs("Cuotar1")
                Basei2 = rs("Basei2")
                Porcei2 = rs("Porcei2")
                Cuotai2 = rs("Cuotai2")
                Porcer2 = rs("Porcer2")
                Cuotar2 = rs("Cuotar2")
                Basei3 = rs("Basei3")
                Porcei3 = rs("Porcei3")
                Cuotai3 = rs("Cuotai3")
                Porcer3 = rs("Porcer3")
                Cuotar3 = rs("Cuotar3")
                Basei4 = rs("Basei4")
                Porcei4 = rs("Porcei4")
                Cuotai4 = rs("Cuotai4")
                Porcer4 = rs("Porcer4")
                Cuotar4 = rs("Cuotar4")
                Base_exenta = rs("Base_exenta")
                Total_factura = rs("Total_factura")
                OPERADOR = rs("OPERADOR")
                CP = rs("CP")
                cr = rs("cr")
                co = rs("co")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsContabilidad_IVARE)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(Nº_Apunte) FROM IVARE"
        Dim oC As New clsContabilidad
        Set rs = oC.datos_bd_contabilidad(consulta)
        Set oC = Nothing
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                Nº_Apunte = 1
        Else
                Nº_Apunte = rs.Fields(0) + 1
        End If
        Nº_Apunte = Format(Nº_Apunte, "000000")
        Set rs = Nothing
End Function

Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        Registro = Nº_Apunte
        consulta = "INSERT INTO IVARE " & _
                   "  VALUES (" & _
                        "'" & Nº_Apunte & "'" & "," & "'" & Fecha_contable & "'" & "," & "'" & Nº_factura & "'" & "," & _
                        "'" & Registro & "'" & "," & "'" & Cuenta & "'" & "," & "'" & Titulo_c & "'" & "," & _
                        "'" & Nif & "'" & "," & "'" & Fecha_factura & "'" & "," & "'" & Basei1 & "'" & "," & _
                        "'" & Porcei1 & "'" & "," & "'" & Cuotai1 & "'" & "," & "'" & Porcer1 & "'" & "," & _
                        "'" & Cuotar1 & "'" & "," & "'" & Basei2 & "'" & "," & "'" & Porcei2 & "'" & "," & _
                        "'" & Cuotai2 & "'" & "," & "'" & Porcer2 & "'" & "," & "'" & Cuotar2 & "'" & "," & _
                        "'" & Basei3 & "'" & "," & "'" & Porcei3 & "'" & "," & "'" & Cuotai3 & "'" & "," & _
                        "'" & Porcer3 & "'" & "," & "'" & Cuotar3 & "'" & "," & "'" & Basei4 & "'" & "," & _
                        "'" & Porcei4 & "'" & "," & "'" & Cuotai4 & "'" & "," & "'" & Porcer4 & "'" & "," & _
                        "'" & Cuotar4 & "'" & "," & "'" & Base_exenta & "'" & "," & "'" & Total_factura & "'" & "," & _
                        "'" & OPERADOR & "'" & "," & "'" & CP & "'" & "," & "'" & cr & "'" & "," & _
                        "'" & co & "'" & _
                ")"
        Dim oC As New clsContabilidad
        oC.execute_bd consulta
        Set oC = Nothing
        Insertar = Nº_Apunte
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsContabilidad_IVARE)", vbCritical, Err.Description
End Function
