VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad_Cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsContabilidad_Cuentas
'***************************************************************
Private Codigo_cuenta As String
Private Titulo_cuenta As String
Private Nif_cuenta As String
Private Recargo_cuenta As String
Private Clave_via As String
Private Nombre_via As String
Private Nº_via As String
Private Cod_postal As String
Private Municipio As String
Private Desglose As String
Private balance As String
Private Debe_apertura As String
Private Haber_apertura As String
Private Debe_enero As String
Private Haber_enero As String
Private Debe_febrero As String
Private Haber_febrero As String
Private Debe_marzo As String
Private Haber_marzo As String
Private Debe_abril As String
Private Haber_abril As String
Private Debe_mayo As String
Private Haber_mayo As String
Private Debe_junio As String
Private Haber_junio As String
Private Debe_julio As String
Private Haber_julio As String
Private Debe_agosto As String
Private Haber_agosto As String
Private Debe_septiembre As String
Private Haber_septiembre As String
Private Debe_octubre As String
Private Haber_octubre As String
Private Debe_noviembre As String
Private Haber_noviembre As String
Private Debe_diciembre As String
Private Haber_diciembre As String
Private el As String
Private ec As String
Private dc As String
Private faclib As Long
Private libro As Long
Private metal As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsContabilidad_Cuentas
'***************************************************************
Public Property Let setCODIGO_CUENTA(ByVal dato As String)
        Codigo_cuenta = dato
End Property
Public Property Let setTITULO_CUENTA(ByVal dato As String)
        Titulo_cuenta = dato
End Property
Public Property Let setNIF_CUENTA(ByVal dato As String)
        Nif_cuenta = dato
End Property
Public Property Let setRECARGO_CUENTA(ByVal dato As String)
        Recargo_cuenta = dato
End Property
Public Property Let setCLAVE_VIA(ByVal dato As String)
        Clave_via = dato
End Property
Public Property Let setNOMBRE_VIA(ByVal dato As String)
        Nombre_via = dato
End Property
Public Property Let setNº_VIA(ByVal dato As String)
        Nº_via = dato
End Property
Public Property Let setCOD_POSTAL(ByVal dato As String)
        Cod_postal = dato
End Property
Public Property Let setMUNICIPIO(ByVal dato As String)
        Municipio = dato
End Property
Public Property Let setDESGLOSE(ByVal dato As String)
        Desglose = dato
End Property
Public Property Let setBALANCE(ByVal dato As String)
        balance = dato
End Property
Public Property Let setDEBE_APERTURA(ByVal dato As String)
        Debe_apertura = dato
End Property
Public Property Let setHABER_APERTURA(ByVal dato As String)
        Haber_apertura = dato
End Property
Public Property Let setDEBE_ENERO(ByVal dato As String)
        Debe_enero = dato
End Property
Public Property Let setHABER_ENERO(ByVal dato As String)
        Haber_enero = dato
End Property
Public Property Let setDEBE_FEBRERO(ByVal dato As String)
        Debe_febrero = dato
End Property
Public Property Let setHABER_FEBRERO(ByVal dato As String)
        Haber_febrero = dato
End Property
Public Property Let setDEBE_MARZO(ByVal dato As String)
        Debe_marzo = dato
End Property
Public Property Let setHABER_MARZO(ByVal dato As String)
        Haber_marzo = dato
End Property
Public Property Let setDEBE_ABRIL(ByVal dato As String)
        Debe_abril = dato
End Property
Public Property Let setHABER_ABRIL(ByVal dato As String)
        Haber_abril = dato
End Property
Public Property Let setDEBE_MAYO(ByVal dato As String)
        Debe_mayo = dato
End Property
Public Property Let setHABER_MAYO(ByVal dato As String)
        Haber_mayo = dato
End Property
Public Property Let setDEBE_JUNIO(ByVal dato As String)
        Debe_junio = dato
End Property
Public Property Let setHABER_JUNIO(ByVal dato As String)
        Haber_junio = dato
End Property
Public Property Let setDEBE_JULIO(ByVal dato As String)
        Debe_julio = dato
End Property
Public Property Let setHABER_JULIO(ByVal dato As String)
        Haber_julio = dato
End Property
Public Property Let setDEBE_AGOSTO(ByVal dato As String)
        Debe_agosto = dato
End Property
Public Property Let setHABER_AGOSTO(ByVal dato As String)
        Haber_agosto = dato
End Property
Public Property Let setDEBE_SEPTIEMBRE(ByVal dato As String)
        Debe_septiembre = dato
End Property
Public Property Let setHABER_SEPTIEMBRE(ByVal dato As String)
        Haber_septiembre = dato
End Property
Public Property Let setDEBE_OCTUBRE(ByVal dato As String)
        Debe_octubre = dato
End Property
Public Property Let setHABER_OCTUBRE(ByVal dato As String)
        Haber_octubre = dato
End Property
Public Property Let setDEBE_NOVIEMBRE(ByVal dato As String)
        Debe_noviembre = dato
End Property
Public Property Let setHABER_NOVIEMBRE(ByVal dato As String)
        Haber_noviembre = dato
End Property
Public Property Let setDEBE_DICIEMBRE(ByVal dato As String)
        Debe_diciembre = dato
End Property
Public Property Let setHABER_DICIEMBRE(ByVal dato As String)
        Haber_diciembre = dato
End Property
Public Property Let setEL(ByVal dato As String)
        el = dato
End Property
Public Property Let setEC(ByVal dato As String)
        ec = dato
End Property
Public Property Let setDC(ByVal dato As String)
        dc = dato
End Property
Public Property Let setFACLIB(ByVal dato As Long)
        faclib = dato
End Property
Public Property Let setLIBRO(ByVal dato As Long)
        libro = dato
End Property
Public Property Let setMETAL(ByVal dato As String)
        metal = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsContabilidad_Cuentas
'***************************************************************
Public Property Get getCODIGO_CUENTA() As String
        getCODIGO_CUENTA = Codigo_cuenta
End Property
Public Property Get getTITULO_CUENTA() As String
        getTITULO_CUENTA = Titulo_cuenta
End Property
Public Property Get getNIF_CUENTA() As String
        getNIF_CUENTA = Nif_cuenta
End Property
Public Property Get getRECARGO_CUENTA() As String
        getRECARGO_CUENTA = Recargo_cuenta
End Property
Public Property Get getCLAVE_VIA() As String
        getCLAVE_VIA = Clave_via
End Property
Public Property Get getNOMBRE_VIA() As String
        getNOMBRE_VIA = Nombre_via
End Property
Public Property Get getNº_VIA() As String
        getNº_VIA = Nº_via
End Property
Public Property Get getCOD_POSTAL() As String
        getCOD_POSTAL = Cod_postal
End Property
Public Property Get getMUNICIPIO() As String
        getMUNICIPIO = Municipio
End Property
Public Property Get getDESGLOSE() As String
        getDESGLOSE = Desglose
End Property
Public Property Get getBALANCE() As String
        getBALANCE = balance
End Property
Public Property Get getDEBE_APERTURA() As String
        getDEBE_APERTURA = Debe_apertura
End Property
Public Property Get getHABER_APERTURA() As String
        getHABER_APERTURA = Haber_apertura
End Property
Public Property Get getDEBE_ENERO() As String
        getDEBE_ENERO = Debe_enero
End Property
Public Property Get getHABER_ENERO() As String
        getHABER_ENERO = Haber_enero
End Property
Public Property Get getDEBE_FEBRERO() As String
        getDEBE_FEBRERO = Debe_febrero
End Property
Public Property Get getHABER_FEBRERO() As String
        getHABER_FEBRERO = Haber_febrero
End Property
Public Property Get getDEBE_MARZO() As String
        getDEBE_MARZO = Debe_marzo
End Property
Public Property Get getHABER_MARZO() As String
        getHABER_MARZO = Haber_marzo
End Property
Public Property Get getDEBE_ABRIL() As String
        getDEBE_ABRIL = Debe_abril
End Property
Public Property Get getHABER_ABRIL() As String
        getHABER_ABRIL = Haber_abril
End Property
Public Property Get getDEBE_MAYO() As String
        getDEBE_MAYO = Debe_mayo
End Property
Public Property Get getHABER_MAYO() As String
        getHABER_MAYO = Haber_mayo
End Property
Public Property Get getDEBE_JUNIO() As String
        getDEBE_JUNIO = Debe_junio
End Property
Public Property Get getHABER_JUNIO() As String
        getHABER_JUNIO = Haber_junio
End Property
Public Property Get getDEBE_JULIO() As String
        getDEBE_JULIO = Debe_julio
End Property
Public Property Get getHABER_JULIO() As String
        getHABER_JULIO = Haber_julio
End Property
Public Property Get getDEBE_AGOSTO() As String
        getDEBE_AGOSTO = Debe_agosto
End Property
Public Property Get getHABER_AGOSTO() As String
        getHABER_AGOSTO = Haber_agosto
End Property
Public Property Get getDEBE_SEPTIEMBRE() As String
        getDEBE_SEPTIEMBRE = Debe_septiembre
End Property
Public Property Get getHABER_SEPTIEMBRE() As String
        getHABER_SEPTIEMBRE = Haber_septiembre
End Property
Public Property Get getDEBE_OCTUBRE() As String
        getDEBE_OCTUBRE = Debe_octubre
End Property
Public Property Get getHABER_OCTUBRE() As String
        getHABER_OCTUBRE = Haber_octubre
End Property
Public Property Get getDEBE_NOVIEMBRE() As String
        getDEBE_NOVIEMBRE = Debe_noviembre
End Property
Public Property Get getHABER_NOVIEMBRE() As String
        getHABER_NOVIEMBRE = Haber_noviembre
End Property
Public Property Get getDEBE_DICIEMBRE() As String
        getDEBE_DICIEMBRE = Debe_diciembre
End Property
Public Property Get getHABER_DICIEMBRE() As String
        getHABER_DICIEMBRE = Haber_diciembre
End Property
Public Property Get getEL() As String
        getEL = el
End Property
Public Property Get getEC() As String
        getEC = ec
End Property
Public Property Get getDC() As String
        getDC = dc
End Property
Public Property Get getFACLIB() As Long
        getFACLIB = faclib
End Property
Public Property Get getLIBRO() As Long
        getLIBRO = libro
End Property
Public Property Get getMETAL() As String
        getMETAL = metal
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsContabilidad_Cuentas
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CUENTAS WHERE Codigo_cuenta = '" & ID & "'"
        Dim oC As New clsContabilidad
        Set rs = oC.datos_bd_contabilidad(consulta)
        Set oC = Nothing
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                Codigo_cuenta = rs("Codigo_cuenta")
                If Not IsNull(rs("Titulo_cuenta")) Then
                    Titulo_cuenta = rs("Titulo_cuenta")
                End If
                If Not IsNull(rs("Nif_cuenta")) Then
                    Nif_cuenta = rs("Nif_cuenta")
                End If
                If Not IsNull(rs("Recargo_cuenta")) Then
                    Recargo_cuenta = rs("Recargo_cuenta")
                End If
                If Not IsNull(rs("Clave_via")) Then
                    Clave_via = rs("Clave_via")
                End If
                If Not IsNull(rs("Nombre_via")) Then
                    Nombre_via = rs("Nombre_via")
                End If
                If Not IsNull(rs("Nº_via")) Then
                    Nº_via = rs("Nº_via")
                End If
                If Not IsNull(rs("Cod_postal")) Then
                    Cod_postal = rs("Cod_postal")
                End If
                If Not IsNull(rs("Cod_postal")) Then
                    Municipio = rs("Cod_postal")
                End If
                If Not IsNull(rs("Desglose")) Then
                    Desglose = rs("Desglose")
                End If
                If Not IsNull(rs("balance")) Then
                    balance = rs("balance")
                End If
                Debe_apertura = rs("Debe_apertura")
                Haber_apertura = rs("Haber_apertura")
                Debe_enero = rs("Debe_enero")
                Haber_enero = rs("Haber_enero")
                Debe_febrero = rs("Debe_febrero")
                Haber_febrero = rs("Haber_febrero")
                Debe_marzo = rs("Debe_marzo")
                Haber_marzo = rs("Haber_marzo")
                Debe_abril = rs("Debe_abril")
                Haber_abril = rs("Haber_abril")
                Debe_mayo = rs("Debe_mayo")
                Haber_mayo = rs("Haber_mayo")
                Debe_junio = rs("Debe_junio")
                Haber_junio = rs("Haber_junio")
                Debe_julio = rs("Debe_julio")
                Haber_julio = rs("Haber_julio")
                Debe_agosto = rs("Debe_agosto")
                Haber_agosto = rs("Haber_agosto")
                Debe_septiembre = rs("Debe_septiembre")
                Haber_septiembre = rs("Haber_septiembre")
                Debe_octubre = rs("Debe_octubre")
                Haber_octubre = rs("Haber_octubre")
                Debe_noviembre = rs("Debe_noviembre")
                Haber_noviembre = rs("Haber_noviembre")
                Debe_diciembre = rs("Debe_diciembre")
                Haber_diciembre = rs("Haber_diciembre")
                If Not IsNull(rs("el")) Then
                    el = rs("el")
                End If
                If Not IsNull(rs("ec")) Then
                    ec = rs("ec")
                End If
                If Not IsNull(rs("dc")) Then
                    dc = rs("dc")
                End If
                If Not IsNull(rs("faclib")) Then
                    faclib = rs("faclib")
                End If
                If Not IsNull(rs("libro")) Then
                    libro = rs("libro")
                End If
                metal = rs("metal")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsContabilidad_Cuentas)", vbCritical, Err.Description
End Function
Public Function Carga_rs(ID As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CUENTAS WHERE Codigo_cuenta = '" & ID & "'"
        Dim oC As New clsContabilidad
        Set Carga_rs = oC.datos_bd_contabilidad(consulta)
        Set oC = Nothing
End Function
Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "INSERT INTO CUENTAS " & _
                   "  VALUES (" & _
                        "'" & Codigo_cuenta & "'" & "," & "'" & Titulo_cuenta & "'" & "," & "'" & Nif_cuenta & "'" & "," & _
                        "'" & Recargo_cuenta & "'" & "," & "'" & Clave_via & "'" & "," & "'" & Nombre_via & "'" & "," & _
                        "'" & Nº_via & "'" & "," & "'" & Cod_postal & "'" & "," & "'" & Municipio & "'" & "," & _
                        "'" & Desglose & "'" & "," & "'" & balance & "'" & "," & "'0" & "'" & "," & _
                        "'0" & "'" & "," & "'0" & "'" & "," & "'0" & "'" & "," & _
                        "'0" & "'" & "," & "'0" & "'" & "," & "'0" & "'" & "," & _
                        "'0" & "'" & "," & "'0" & "'" & "," & "'0" & "'" & "," & _
                        "'0" & "'" & "," & "'0" & "'" & "," & "'0" & "'" & "," & _
                        "'0" & "'" & "," & "'0" & "'" & "," & "'0" & "'" & "," & _
                        "'0" & "'" & "," & "'0" & "'" & "," & "'0" & "'" & "," & _
                        "'0" & "'" & "," & "'0" & "'" & "," & "'0" & "'" & "," & _
                        "'0" & "'" & "," & "'0" & "'" & "," & "'0" & "'" & "," & _
                        "'0" & "'" & "," & "'" & el & "'" & "," & "'" & ec & "'" & "," & _
                        "'" & dc & "'" & "," & faclib & "," & libro & "," & _
                        metal & _
                ")"
        Dim oC As New clsContabilidad
        oC.execute_bd consulta
        Set oC = Nothing
        Insertar = Codigo_cuenta
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsContabilidad_Cuentas)", vbCritical, Err.Description
End Function
