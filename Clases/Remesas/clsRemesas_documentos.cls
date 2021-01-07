VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemesas_documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSREMESAS_DOCUMENTOS
'***************************************************************
Private REMESA_ID As Long
Private FACTURA_ID As Long


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSREMESAS_DOCUMENTOS
'***************************************************************
Public Property Let setREMESA_ID(ByVal dato As Long)
        REMESA_ID = dato
End Property
Public Property Let setFACTURA_ID(ByVal dato As Long)
        FACTURA_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSREMESAS_DOCUMENTOS
'***************************************************************
Public Property Get getREMESA_ID() As Long
        getREMESA_ID = REMESA_ID
End Property
Public Property Get getFACTURA_ID() As Long
        getFACTURA_ID = FACTURA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSREMESAS_DOCUMENTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REMESAS_DOCUMENTOS WHERE FACTURA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                REMESA_ID = rs("REMESA_ID")
                FACTURA_ID = rs("FACTURA_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRemesas_documentos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO REMESAS_DOCUMENTOS " & _
                   "  VALUES (" & _
                        REMESA_ID & "," & FACTURA_ID & _
                ")"
        execute_bd consulta
        Insertar = REMESA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRemesas_documentos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE REMESAS_DOCUMENTOS SET " & _
                        " FACTURA_ID = " & FACTURA_ID & "" & _
                " WHERE REMESA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRemesas_documentos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDORES_FACTURAS " & _
                   "   SET F_PAGO = '0000-00-00', REMESA_ID = 0 " & _
                   " WHERE REMESA_ID = " & ID
        execute_bd consulta
        consulta = "DELETE FROM REMESAS_DOCUMENTOS " & _
                "    WHERE REMESA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRemesas_documentos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REMESAS_DOCUMENTOS ORDER BY REMESA_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSREMESAS_DOCUMENTOS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

