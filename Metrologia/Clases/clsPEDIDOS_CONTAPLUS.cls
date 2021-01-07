VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPEDIDOS_CONTAPLUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPEDIDOS_CONTAPLUS
'***************************************************************
Private CCOMENT As String
Private CNOMCLI As String
Private DFECPED As String
Private DOCUMENTO_ID As Long
Private NNUMPED As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPEDIDOS_CONTAPLUS
'***************************************************************
Public Property Let setCCOMENT(ByVal dato As String)
        CCOMENT = dato
End Property
Public Property Let setCNOMCLI(ByVal dato As String)
        CNOMCLI = dato
End Property
Public Property Let setDFECPED(ByVal dato As String)
        DFECPED = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setNNUMPED(ByVal dato As Long)
        NNUMPED = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPEDIDOS_CONTAPLUS
'***************************************************************
Public Property Get getCCOMENT() As String
        getCCOMENT = CCOMENT
End Property
Public Property Get getCNOMCLI() As String
        getCNOMCLI = CNOMCLI
End Property
Public Property Get getDFECPED() As String
        getDFECPED = DFECPED
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getNNUMPED() As Long
        getNNUMPED = NNUMPED
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPEDIDOS_CONTAPLUS
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PEDIDOS_CONTAPLUS WHERE NNUMPED = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CCOMENT = rs("CCOMENT")
                CNOMCLI = rs("CNOMCLI")
                DFECPED = rs("DFECPED")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                NNUMPED = rs("NNUMPED")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPEDIDOS_CONTAPLUS)", vbCritical, Err.Description
End Function
Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PEDIDOS_CONTAPLUS " & _
                   "  VALUES (" & _
                        NNUMPED & ",'" & DFECPED & "'" & "," & "'" & CNOMCLI & "'" & "," & "'" & CCOMENT & "'" & "," & _
                        DOCUMENTO_ID & _
                ")"
        execute_bd consulta
        Insertar = NNUMPED
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPEDIDOS_CONTAPLUS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PEDIDOS_CONTAPLUS SET " & _
                        " CNOMCLI = '" & CNOMCLI & "'," & _
                        " DFECPED = '" & DFECPED & "'," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & _
                " WHERE NNUMPED = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPEDIDOS_CONTAPLUS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PEDIDOS_CONTAPLUS " & _
                "    WHERE NNUMPED = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPEDIDOS_CONTAPLUS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PEDIDOS_CONTAPLUS ORDER BY NNUMPED"
        Set Listado = datos_bd(consulta)
End Function
Public Function Ultimo_pedido() As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(NNUMPED) FROM PEDIDOS_CONTAPLUS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Ultimo_pedido = 0
    Else
        Ultimo_pedido = rs.Fields(0)
    End If
    Set rs = Nothing
End Function

