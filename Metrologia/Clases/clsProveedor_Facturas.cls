VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedor_Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROVEEDOR_FACTURAS
'***************************************************************
Private ID As Long
Private PROVEEDOR_ID As Long
Private fecha As String
Private numero As String
Private BI As String
Private iva As String
Private total As String
Private FORMAPAGO As Integer
Private Documento As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROVEEDOR_FACTURAS
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setNUMERO(ByVal dato As String)
        numero = dato
End Property
Public Property Let setBI(ByVal dato As String)
        BI = dato
End Property
Public Property Let setIVA(ByVal dato As String)
        iva = dato
End Property
Public Property Let setTOTAL(ByVal dato As String)
        total = dato
End Property
Public Property Let setFORMAPAGO(ByVal dato As Integer)
        FORMAPAGO = dato
End Property
Public Property Let setDOCUMENTO(ByVal dato As String)
        Documento = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROVEEDOR_FACTURAS
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getNUMERO() As String
        getNUMERO = numero
End Property
Public Property Get getBI() As String
        getBI = BI
End Property
Public Property Get getIVA() As String
        getIVA = iva
End Property
Public Property Get getTOTAL() As String
        getTOTAL = total
End Property
Public Property Get getFORMAPAGO() As Integer
        getFORMAPAGO = FORMAPAGO
End Property
Public Property Get getDOCUMENTO() As String
        getDOCUMENTO = Documento
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROVEEDOR_FACTURAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDOR_FACTURAS WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                fecha = rs("FECHA")
                numero = rs("NUMERO")
                BI = rs("BI")
                iva = rs("IVA")
                total = rs("TOTAL")
                FORMAPAGO = rs("FORMAPAGO")
                Documento = rs("DOCUMENTO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPROVEEDOR_FACTURAS)", vbCritical, Err.DescrBItion
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM PROVEEDOR_FACTURAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PROVEEDOR_FACTURAS " & _
                   "  VALUES (" & _
                        ID & "," & PROVEEDOR_ID & "," & "'" & Format(fecha, "yyyy-mm-dd") & "'" & "," & _
                        "'" & numero & "'" & "," & "'" & BI & "'" & "," & "'" & iva & "'" & "," & _
                        "'" & total & "'," & FORMAPAGO & ",'" & Documento & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPROVEEDOR_FACTURAS)", vbCritical, Err.DescrBItion
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDOR_FACTURAS SET " & _
                        " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " NUMERO = '" & numero & "'," & _
                        " BI = '" & BI & "'," & _
                        " IVA = '" & iva & "'," & _
                        " DOCUMENTO = '" & Documento & "'," & _
                        " FORMAPAGO = " & FORMAPAGO & "," & _
                        " TOTAL = '" & total & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPROVEEDOR_FACTURAS)", vbCritical, Err.DescrBItion
End Function
Public Function informar_documento(ID As Long, Documento As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDOR_FACTURAS SET " & _
                        " DOCUMENTO = '" & Documento & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        informar_documento = True
        Exit Function
fallo:
        informar_documento = False
        MsgBox "Error al informar_documento (clsPROVEEDOR_FACTURAS)", vbCritical, Err.DescrBItion
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROVEEDOR_FACTURAS " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPROVEEDOR_FACTURAS)", vbCritical, Err.DescrBItion
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT FECHA,NUMERO,BI,IVA,TOTAL,FP.NOMBRE,ID,DOCUMENTO " & _
                   "  FROM PROVEEDOR_FACTURAS PF, FORMA_PAGO FP " & _
                   " WHERE PROVEEDOR_ID = " & ID & _
                   "   AND PF.FORMAPAGO = FP.ID_FORMA_PAGO " & _
                   " ORDER BY FECHA DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_IVA() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT FECHA,NUMERO,BI,IVA,TOTAL,FP.NOMBRE,ID,DOCUMENTO " & _
                   "  FROM PROVEEDOR_FACTURAS PF, FORMA_PAGO FP " & _
                   " WHERE PF.FORMAPAGO = FP.ID_FORMA_PAGO " & _
                   " ORDER BY FECHA DESC"
        Set Listado_IVA = datos_bd(consulta)
End Function

