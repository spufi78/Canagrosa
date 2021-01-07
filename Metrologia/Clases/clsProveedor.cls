VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE PROVEEDOR    *
'*****************************************************
Private ID_PROVEEDOR As Long
Private NOMBRE As String
Private ACTIVIDAD As String
Private DIRECCION As String
Private MUNICIPIO_ID As Long
Private PROVINCIA_ID As Long
Private CP As Long
Private direccionN As String
Private MUNICIPION_ID As Long
Private PROVINCIAN_ID As Long
Private CPN As Long
Private CIF As String
Private TELEFONO As String
Private movil As String
Private FAX As String
Private EMAIL As String
Private RESPONSABLE As String
Private OBSERVACIONES As String
Private FORMA_PAGO As Integer
Private BANCO As String
Private ccc As String
'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE PROVEEDOR    *
'*****************************************************
Public Property Let setID_PROVEEDOR(ByVal dato As Long)
    ID_PROVEEDOR = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setACTIVIDAD(ByVal dato As String)
    ACTIVIDAD = Trim(dato)
End Property
Public Property Let setDIRECCION(ByVal dato As String)
    DIRECCION = Trim(dato)
End Property
Public Property Let setCP(ByVal dato As Long)
    CP = dato
End Property
Public Property Let setDIRECCIONN(ByVal dato As String)
    direccionN = Trim(dato)
End Property
Public Property Let setCPN(ByVal dato As Long)
    CPN = dato
End Property
Public Property Let setCIF(ByVal dato As String)
    CIF = Trim(dato)
End Property
Public Property Let setTELEFONO(ByVal dato As String)
    TELEFONO = Trim(dato)
End Property
Public Property Let setMOVIL(ByVal dato As String)
    movil = Trim(dato)
End Property
Public Property Let setFAX(ByVal dato As String)
    FAX = Trim(dato)
End Property
Public Property Let setEMAIL(ByVal dato As String)
    EMAIL = Trim(dato)
End Property
Public Property Let setRESPONSABLE(ByVal dato As String)
    RESPONSABLE = Trim(dato)
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
    MUNICIPIO_ID = dato
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
    PROVINCIA_ID = dato
End Property
Public Property Let setMUNICIPION_ID(ByVal dato As Long)
    MUNICIPION_ID = dato
End Property
Public Property Let setPROVINCIAN_ID(ByVal dato As Long)
    PROVINCIAN_ID = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setFORMA_PAGO(ByVal dato As Integer)
    FORMA_PAGO = dato
End Property
Public Property Let setBANCO(ByVal dato As String)
    BANCO = Trim(dato)
End Property
Public Property Let setCCC(ByVal dato As String)
    ccc = Trim(dato)
End Property

'*****************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE PROVEEDOR      *
'*****************************************************
Public Property Get getID_PROVEEDOR() As Long
    getID_PROVEEDOR = ID_PROVEEDOR
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getACTIVIDAD() As String
    getACTIVIDAD = ACTIVIDAD
End Property
Public Property Get getDIRECCION() As String
    getDIRECCION = DIRECCION
End Property
Public Property Get getCP() As Long
    getCP = CP
End Property
Public Property Get getDIRECCIONN() As String
    getDIRECCIONN = direccionN
End Property
Public Property Get getCPN() As Long
    getCPN = CPN
End Property
Public Property Get getCIF() As String
    getCIF = CIF
End Property
Public Property Get getTELEFONO() As String
    getTELEFONO = TELEFONO
End Property
Public Property Get getMOVIL() As String
    getMOVIL = movil
End Property
Public Property Get getFAX() As String
    getFAX = FAX
End Property
Public Property Get getEMAIL() As String
    getEMAIL = EMAIL
End Property
Public Property Get getRESPONSABLE() As String
    getRESPONSABLE = RESPONSABLE
End Property
Public Property Get getMUNICIPIO_ID() As Long
    getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getPROVINCIA_ID() As Long
    getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getMUNICIPION_ID() As Long
    getMUNICIPION_ID = MUNICIPION_ID
End Property
Public Property Get getPROVINCIAN_ID() As Long
    getPROVINCIAN_ID = PROVINCIAN_ID
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFORMA_PAGO() As Integer
    getFORMA_PAGO = FORMA_PAGO
End Property
Public Property Get getBANCO() As String
    getBANCO = BANCO
End Property
Public Property Get getCCC() As String
    getCCC = ccc
End Property

'*****************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE MUESTRA  *
'*****************************************************
Public Function Carga(ID_PROVEEDOR As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES WHERE ID_PROVEEDOR=" & ID_PROVEEDOR
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            ID_PROVEEDOR = rs.Fields("ID_PROVEEDOR")
            NOMBRE = rs.Fields("NOMBRE") & ""
            ACTIVIDAD = rs.Fields("ACTIVIDAD") & ""
            DIRECCION = rs.Fields("DIRECCION") & ""
            If Not IsNull(rs.Fields("CP")) Then
                CP = rs.Fields("CP")
            Else
                CP = 0
            End If
            PROVINCIA_ID = rs.Fields("PROVINCIA_ID")
            MUNICIPIO_ID = rs.Fields("MUNICIPIO_ID")
            direccionN = rs.Fields("DIRECCIONN") & ""
            If Not IsNull(rs.Fields("CPN")) Then
                CPN = rs.Fields("CPN")
            Else
                CPN = 0
            End If
            PROVINCIAN_ID = rs.Fields("PROVINCIAN_ID")
            MUNICIPION_ID = rs.Fields("MUNICIPION_ID")
            CIF = rs.Fields("CIF") & ""
            TELEFONO = rs.Fields("TELEFONO") & ""
            movil = rs.Fields("MOVIL") & ""
            FAX = rs.Fields("FAX") & ""
            EMAIL = rs.Fields("EMAIL") & ""
            RESPONSABLE = rs.Fields("RESPONSABLE") & ""
            OBSERVACIONES = rs.Fields("OBSERVACIONES") & ""
            FORMA_PAGO = rs.Fields("FORMA_PAGO")
            BANCO = rs.Fields("BANCO") & ""
            ccc = rs.Fields("CCC") & ""
            Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos del proveedor (Carga)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
    Dim consulta As String
    On Error GoTo fallo:
    If ID_PROVEEDOR = 0 Then
        Me.CrearID_PROVEEDOR
    End If
    consulta = "INSERT INTO PROVEEDORES " & _
               "values(" & _
               ID_PROVEEDOR & ",'" & NOMBRE & "','" & ACTIVIDAD & "'," & _
               "'" & DIRECCION & "'," & CP & "," & PROVINCIA_ID & "," & MUNICIPIO_ID & "," & _
               "'" & direccionN & "'," & CPN & "," & PROVINCIAN_ID & "," & MUNICIPION_ID & "," & _
               "'" & CIF & "','" & _
               TELEFONO & "','" & movil & "','" & FAX & "','" & EMAIL & "','" & _
               OBSERVACIONES & "','" & RESPONSABLE & "'," & FORMA_PAGO & ",'" & _
               BANCO & "','" & ccc & "')"
    execute_bd consulta
    Insertar = ID_PROVEEDOR
'    MsgBox "Proveedor almacenado correctamente. Código : " & ID_PROVEEDOR, vbInformation, App.Title
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al guardar el proveedor en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Sub CrearID_PROVEEDOR()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PROVEEDOR) FROM PROVEEDORES"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PROVEEDOR = 1
    Else
        ID_PROVEEDOR = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Modificar() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE PROVEEDORES SET " & _
               "NOMBRE = '" & NOMBRE & "',ACTIVIDAD='" & ACTIVIDAD & "'" & _
               ",DIRECCION = '" & DIRECCION & "',CP = " & CP & _
               ",DIRECCIONN = '" & direccionN & "',CPN = " & CPN & _
               ",CIF = '" & CIF & "',TELEFONO = '" & TELEFONO & "'" & _
               ",FAX = '" & FAX & "',RESPONSABLE = '" & RESPONSABLE & "'" & _
               ",OBSERVACIONES = '" & OBSERVACIONES & "'" & _
               ",PROVINCIA_ID = " & PROVINCIA_ID & _
               ",MUNICIPIO_ID = " & MUNICIPIO_ID & _
               ",PROVINCIAN_ID = " & PROVINCIAN_ID & _
               ",MUNICIPION_ID = " & MUNICIPION_ID & _
               ",MOVIL = '" & movil & "'" & _
               ",EMAIL = '" & EMAIL & "'" & _
               ",FORMA_PAGO = " & FORMA_PAGO & _
               ",BANCO = '" & BANCO & "'" & _
               ",CCC = '" & ccc & "'" & _
               " WHERE " & _
               " ID_PROVEEDOR = " & ID_PROVEEDOR
               execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el proveedor en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Eliminar(ID As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "DELETE FROM PROVEEDORES " & _
               " WHERE " & _
               " ID_PROVEEDOR = " & ID
               execute_bd consulta
    Eliminar = True
    MsgBox "Proveedor eliminado correctamente.", vbInformation, App.Title
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el proveedor de la bd : " & Err.Description, vbCritical, Err.Number
End Function

Public Function Listado(pNOMBRE As String, pCIF As String, pTELEFONO As String, pFAX As String) As ADODB.Recordset
    Dim consulta As String
    Dim s As String
    If Trim(pNOMBRE) <> "" Then
        s = s & " AND NOMBRE LIKE '%" & pNOMBRE & "%'"
    End If
    If Trim(pCIF) <> "" Then
        s = s & " AND CIF LIKE '%" & pCIF & "%'"
    End If
    If Trim(pTELEFONO) <> "" Then
        s = s & " AND TELEFONO LIKE '%" & pTELEFONO & "%'"
    End If
    If Trim(pFAX) <> "" Then
        s = s & " AND FAX LIKE '%" & pFAX & "%'"
    End If
    consulta = "SELECT * FROM PROVEEDORES " & _
               " WHERE ID_PROVEEDOR <> 0 " & _
               s & _
               " order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Busqueda(pNOMBRE As String, pCIF As String, pTELEFONO As String, pFAX As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_PROVEEDOR,NOMBRE,DIRECCION,TELEFONO,FAX " & _
               "  FROM PROVEEDORES " & _
               " WHERE NOMBRE LIKE '%" & pNOMBRE & "%'" & _
               "   AND CIF    LIKE '%" & pCIF & "%'" & _
               "   AND TELEFONO LIKE '%" & pTELEFONO & "%'" & _
               "   AND FAX    LIKE '%" & pFAX & "%'" & _
               " ORDER BY NOMBRE"
    Set Busqueda = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_PROVEEDOR, NOMBRE FROM PROVEEDORES WHERE ID_PROVEEDOR <> 0 " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_Combo_articulo(ARTICULO As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.ID_PROVEEDOR, A.NOMBRE " & _
               "  FROM ARTICULOS_PROVEEDORES B,PROVEEDORES A " & _
               " WHERE B.ARTICULO_ID = " & ARTICULO & _
               "   AND B.PROVEEDOR_ID = A.ID_PROVEEDOR " & _
               " order by A.NOMBRE"
    Set Listado_Combo_articulo = datos_bd(consulta)
End Function
Public Function llenar_combo(combo As miComboBCA, FK As Long, FORMULARIO As Form, FILTRO As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "PROVEEDORES"
        .setDESCRIPCION = "Proveedores"
        .setPK = "ID_PROVEEDOR"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        .setFILTRO = FILTRO
        Set .FORMULARIO = FORMULARIO
    End With
End Function

