VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad_Movimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsContabilidad_Movimientos
'***************************************************************
Private Apunte As String
Private Apunte_diario As String
Private fecha As String
Private Documento As String
Private Cuenta As String
Private Descripcion_apunte As String
Private Tipo_apunte As String
Private Importe_apunte As String
Private contrapartida As String
Private CC As String
Private OPERADOR As String
Private metalico As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsContabilidad_Movimientos
'***************************************************************
Public Property Let setAPUNTE(ByVal dato As String)
        Apunte = dato
End Property
Public Property Let setAPUNTE_DIARIO(ByVal dato As String)
        Apunte_diario = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setDOCUMENTO(ByVal dato As String)
        Documento = dato
End Property
Public Property Let setCUENTA(ByVal dato As String)
        Cuenta = dato
End Property
Public Property Let setDESCRIPCION_APUNTE(ByVal dato As String)
        Descripcion_apunte = dato
End Property
Public Property Let setTIPO_APUNTE(ByVal dato As String)
        Tipo_apunte = dato
End Property
Public Property Let setIMPORTE_APUNTE(ByVal dato As String)
        Importe_apunte = dato
End Property
Public Property Let setCONTRAPARTIDA(ByVal dato As String)
        contrapartida = dato
End Property
Public Property Let setCC(ByVal dato As String)
        CC = dato
End Property
Public Property Let setOPERADOR(ByVal dato As String)
        OPERADOR = dato
End Property
Public Property Let setMETALICO(ByVal dato As String)
        metalico = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsContabilidad_Movimientos
'***************************************************************
Public Property Get getAPUNTE() As String
        getAPUNTE = Apunte
End Property
Public Property Get getAPUNTE_DIARIO() As String
        getAPUNTE_DIARIO = Apunte_diario
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getDOCUMENTO() As String
        getDOCUMENTO = Documento
End Property
Public Property Get getCUENTA() As String
        getCUENTA = Cuenta
End Property
Public Property Get getDESCRIPCION_APUNTE() As String
        getDESCRIPCION_APUNTE = Descripcion_apunte
End Property
Public Property Get getTIPO_APUNTE() As String
        getTIPO_APUNTE = Tipo_apunte
End Property
Public Property Get getIMPORTE_APUNTE() As String
        getIMPORTE_APUNTE = Importe_apunte
End Property
Public Property Get getCONTRAPARTIDA() As String
        getCONTRAPARTIDA = contrapartida
End Property
Public Property Get getCC() As String
        getCC = CC
End Property
Public Property Get getOPERADOR() As String
        getOPERADOR = OPERADOR
End Property
Public Property Get getMETALICO() As String
        getMETALICO = metalico
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsContabilidad_Movimientos
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MOVIMIENTOS WHERE Apunte = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                Apunte = rs("Apunte")
                Apunte_diario = rs("Apunte_diario")
                fecha = rs("Fecha")
                Documento = rs("Documento")
                Cuenta = rs("Cuenta")
                Descripcion_apunte = rs("Descripcion_apunte")
                Tipo_apunte = rs("Tipo_apunte")
                Importe_apunte = rs("Importe_apunte")
                contrapartida = rs("Contrapartida")
                CC = rs("CC")
                OPERADOR = rs("OPERADOR")
                metalico = rs("metalico")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsContabilidad_Movimientos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(Apunte) FROM MOVIMIENTOS"
        Dim oC As New clsContabilidad
        Set rs = oC.datos_bd_contabilidad(consulta)
        Set oC = Nothing
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                Apunte = 1
        Else
                Apunte = rs.Fields(0) + 1
        End If
        Apunte = Format(Apunte, "000000")
        Set rs = Nothing
End Function
Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        Apunte_diario = Apunte
        consulta = "INSERT INTO MOVIMIENTOS " & _
                   "  VALUES (" & _
                        "'" & Apunte & "'" & "," & "'" & Apunte_diario & "'" & "," & "'" & fecha & "'" & "," & _
                        "'" & Documento & "'" & "," & "'" & Cuenta & "'" & "," & "'" & Descripcion_apunte & "'" & "," & _
                        "'" & Tipo_apunte & "'" & "," & "'" & Importe_apunte & "'" & "," & "'" & contrapartida & "'" & "," & _
                        "'" & CC & "'" & "," & "'" & OPERADOR & "'" & "," & "'" & metalico & "'" & _
                ")"
        Dim oC As New clsContabilidad
        oC.execute_bd consulta
        Set oC = Nothing
        Insertar = Apunte
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsContabilidad_Movimientos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MOVIMIENTOS SET " & _
                        " APUNTE_DIARIO = '" & Apunte_diario & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " DOCUMENTO = '" & Documento & "'," & _
                        " CUENTA = '" & Cuenta & "'," & _
                        " DESCRIPCION_APUNTE = '" & Descripcion_apunte & "'," & _
                        " TIPO_APUNTE = '" & Tipo_apunte & "'," & _
                        " IMPORTE_APUNTE = '" & Importe_apunte & "'," & _
                        " CONTRAPARTIDA = '" & contrapartida & "'," & _
                        " CC = '" & CC & "'," & _
                        " OPERADOR = '" & OPERADOR & "'," & _
                        " METALICO = '" & metalico & "'" & _
                " WHERE Apunte = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsContabilidad_Movimientos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MOVIMIENTOS " & _
                "    WHERE Apunte = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsContabilidad_Movimientos)", vbCritical, Err.Description
End Function
Public Function Listado(FECHAS As Boolean, fdesde As String, fhasta As String, Cuenta As String, contrapartida As String, DESCRIPCION As String) As ADODB.Recordset
        Dim consulta As String
        Dim C As String
        C = ""
        If FECHAS = True Then
            C = C & " AND M.FECHA >= #" & Format(fdesde, "mm/dd/yyyy") & "# AND M.FECHA <= #" & Format(fhasta, "mm/dd/yyyy") & "#"
        End If
        If Cuenta <> "" Then
            C = C & " AND M.CUENTA LIKE '%" & Cuenta & "%'"
        End If
        If contrapartida <> "" Then
            C = C & " AND M.CONTRAPARTIDA LIKE '%" & contrapartida & "%'"
        End If
        If DESCRIPCION <> "" Then
            C = C & " AND M.DESCRIPCION_APUNTE LIKE '%" & DESCRIPCION & "%'"
        End If
        consulta = "SELECT M.APUNTE, M.FECHA, M.DOCUMENTO, M.CUENTA,C1.TITULO_CUENTA,M.DESCRIPCION_APUNTE, " & _
                   "       M.IMPORTE_APUNTE,  M.TIPO_APUNTE,M.CONTRAPARTIDA,C2.TITULO_CUENTA " & _
                   "  FROM MOVIMIENTOS M, CUENTAS C1, CUENTAS C2 " & _
                   " WHERE M.CUENTA = C1.CODIGO_CUENTA " & _
                   "   AND M.CONTRAPARTIDA = C2.CODIGO_CUENTA " & _
                   C & _
                   " ORDER BY APUNTE DESC"
        Dim oC As New clsContabilidad
        Set Listado = oC.datos_bd_contabilidad(consulta)
        Set oC = Nothing
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT APUNTE,APUNTE_DIARIO FROM MOVIMIENTOS ORDER BY Apunte_diario"
        Set Listado_Combo = datos_bd(consulta)
End Function
