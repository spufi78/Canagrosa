VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedores_pto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROVEEDORES_PTO
'***************************************************************
Private PROVEEDOR_ID As Long
Private ANO As Long
Private PRESUPUESTO As String
Private REAL As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROVEEDORES_PTO
'***************************************************************
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setANO(ByVal dato As Long)
        ANO = dato
End Property
Public Property Let setPRESUPUESTO(ByVal dato As String)
        PRESUPUESTO = dato
End Property
Public Property Let setREAL(ByVal dato As String)
        REAL = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROVEEDORES_PTO
'***************************************************************
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getANO() As Long
        getANO = ANO
End Property
Public Property Get getPRESUPUESTO() As String
        getPRESUPUESTO = PRESUPUESTO
End Property
Public Property Get getREAL() As String
        getREAL = REAL
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROVEEDORES_PTO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_PTO WHERE PROVEEDOR_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                ANO = rs("ANO")
                PRESUPUESTO = rs("PRESUPUESTO")
                REAL = rs("REAL")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsProveedores_pto)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PROVEEDORES_PTO " & _
                   "  VALUES (" & _
                        PROVEEDOR_ID & "," & ANO & ",'" & PRESUPUESTO & "','" & _
                        REAL & _
                "')"
        execute_bd consulta
        Insertar = PROVEEDOR_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProveedores_pto)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDORES_PTO SET " & _
                        " ANO = " & ANO & "," & _
                        " PRESUPUESTO = '" & PRESUPUESTO & "'," & _
                        " REAL = '" & REAL & "'" & _
                " WHERE PROVEEDOR_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsProveedores_pto)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROVEEDORES_PTO " & _
                "    WHERE PROVEEDOR_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsProveedores_pto)", vbCritical, Err.Description
End Function
Public Function Listado(PROVEEDOR_ID As Long, ANNO As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_PTO " & _
                   " WHERE PROVEEDOR_ID =" & PROVEEDOR_ID & _
                   "   AND ANO = " & ANNO
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPROVEEDORES_PTO
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

