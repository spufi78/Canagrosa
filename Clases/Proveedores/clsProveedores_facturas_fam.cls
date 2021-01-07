VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedores_facturas_fam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROVEEDORES_FACTURAS_FAM
'***************************************************************
Private ID As Long
Private FACTURA_ID As Long
Private FAMILIA_ID As Long
Private SUBCUENTA_ID As Long
Private BI As String
Private IVA_PORCENTAJE As String
Private IVA As String
Private IMPORTE As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROVEEDORES_FACTURAS_FAM
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setFACTURA_ID(ByVal dato As Long)
        FACTURA_ID = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Long)
        FAMILIA_ID = dato
End Property
Public Property Let setSUBCUENTA_ID(ByVal dato As Long)
        SUBCUENTA_ID = dato
End Property
Public Property Let setBI(ByVal dato As String)
        BI = dato
End Property
Public Property Let setIVA_PORCENTAJE(ByVal dato As String)
        IVA_PORCENTAJE = dato
End Property
Public Property Let setIVA(ByVal dato As String)
        IVA = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROVEEDORES_FACTURAS_FAM
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getFACTURA_ID() As Long
        getFACTURA_ID = FACTURA_ID
End Property
Public Property Get getFAMILIA_ID() As Long
        getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getSUBCUENTA_ID() As Long
        getSUBCUENTA_ID = SUBCUENTA_ID
End Property
Public Property Get getBI() As String
        getBI = BI
End Property
Public Property Get getIVA_PORCENTAJE() As String
        getIVA_PORCENTAJE = IVA_PORCENTAJE
End Property
Public Property Get getIVA() As String
        getIVA = IVA
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROVEEDORES_FACTURAS_FAM
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_FACTURAS_FAM WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                FACTURA_ID = rs("FACTURA_ID")
                FAMILIA_ID = rs("FAMILIA_ID")
                SUBCUENTA_ID = rs("SUBCUENTA_ID")
                BI = rs("BI")
                IVA_PORCENTAJE = rs("IVA_PORCENTAJE")
                IVA = rs("IVA")
                IMPORTE = rs("IMPORTE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsProveedores_facturas_fam)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM PROVEEDORES_FACTURAS_FAM"
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
        consulta = "INSERT INTO PROVEEDORES_FACTURAS_FAM " & _
                   "  VALUES (" & _
                        ID & "," & FACTURA_ID & "," & FAMILIA_ID & "," & _
                        SUBCUENTA_ID & ",'" & BI & "','" & IVA_PORCENTAJE & "','" & _
                        IVA & "','" & IMPORTE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProveedores_facturas_fam)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDORES_FACTURAS_FAM SET " & _
                        " FACTURA_ID = " & FACTURA_ID & "," & _
                        " FAMILIA_ID = " & FAMILIA_ID & "," & _
                        " SUBCUENTA_ID = " & SUBCUENTA_ID & "," & _
                        " BI = '" & BI & "'," & _
                        " IVA_PORCENTAJE = '" & IVA_PORCENTAJE & "'," & _
                        " IVA = '" & IVA & "'," & _
                        " IMPORTE = '" & IMPORTE & "' " & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsProveedores_facturas_fam)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROVEEDORES_FACTURAS_FAM " & _
                "    WHERE FACTURA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsProveedores_facturas_fam)", vbCritical, Err.Description
End Function
Public Function Listado(ID_FACTURA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT PFF.ID,F.NOMBRE,PFF.FAMILIA_ID,D.DESCRIPCION,PFF.SUBCUENTA_ID,PFF.BI,PFF.IVA_PORCENTAJE,PFF.IVA,PFF.IMPORTE " & _
                   "  FROM PROVEEDORES_FACTURAS_FAM PFF " & _
                   "  LEFT JOIN FAMILIAS F ON PFF.FAMILIA_ID = F.ID_FAMILIA " & _
                   "  LEFT JOIN decodificadora D ON PFF.SUBCUENTA_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS & _
                   " WHERE PFF.FACTURA_ID = " & ID_FACTURA & _
                   " ORDER BY PFF.ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPROVEEDORES_FACTURAS_FAM
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,FACTURA_ID FROM PROVEEDORES_FACTURAS_FAM ORDER BY FACTURA_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
'JGM-I
'Public Function obtenerIVAMax(ID) As String
'        Dim rs As ADODB.Recordset
'        Dim consulta As String
'        consulta = "SELECT MAX(IVA_PORCENTAJE) FROM PROVEEDORES_FACTURAS_FAM WHERE FACTURA_ID = " & ID
'        Set rs = datos_bd(consulta)
'        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
'                obtenerIVAMax = "0.00"
'        Else
'                obtenerIVAMax = rs(0)
'        End If
'        Set rs = Nothing
'End Function
'JGM-F
