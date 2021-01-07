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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSREMESAS_DOCUMENTOS
'***************************************************************
Private ID As Long
Private REMESA_ID As Long
Private DOCUMENTO_ID As Long
Private VENCIMIENTO As Long
Private CLIENTE_ID As Long
Private FECHA_VENCIMIENTO As String
Private DESCRIPCION As String
Private IMPORTE As String
Private SITUACION As String
Private CONTA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSREMESAS_DOCUMENTOS
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setREMESA_ID(ByVal dato As Long)
        REMESA_ID = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setVENCIMIENTO(ByVal dato As Long)
        VENCIMIENTO = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setFECHA_VENCIMIENTO(ByVal dato As String)
        FECHA_VENCIMIENTO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
Public Property Let setSITUACION(ByVal dato As String)
        SITUACION = dato
End Property
Public Property Let setCONTA(ByVal dato As String)
        CONTA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSREMESAS_DOCUMENTOS
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getREMESA_ID() As Long
        getREMESA_ID = REMESA_ID
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getVENCIMIENTO() As Long
        getVENCIMIENTO = VENCIMIENTO
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getFECHA_VENCIMIENTO() As String
        getFECHA_VENCIMIENTO = FECHA_VENCIMIENTO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
Public Property Get getSITUACION() As String
        getSITUACION = SITUACION
End Property
Public Property Get getCONTA() As String
        getCONTA = CONTA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSREMESAS_DOCUMENTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REMESAS_DOCUMENTOS WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                REMESA_ID = rs("REMESA_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                VENCIMIENTO = rs("VENCIMIENTO")
                CLIENTE_ID = rs("CLIENTE_ID")
                FECHA_VENCIMIENTO = rs("FECHA_VENCIMIENTO")
                DESCRIPCION = rs("DESCRIPCION")
                IMPORTE = rs("IMPORTE")
                SITUACION = rs("SITUACION")
                CONTA = rs("CONTA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRemesas_documentos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM REMESAS_DOCUMENTOS"
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
        If ID = 0 Then
            Me.CrearID
        End If
        consulta = "INSERT INTO REMESAS_DOCUMENTOS " & _
                   "  VALUES (" & _
                        ID & "," & REMESA_ID & "," & DOCUMENTO_ID & "," & _
                        VENCIMIENTO & "," & CLIENTE_ID & "," & "'" & FECHA_VENCIMIENTO & "'" & "," & _
                        "'" & DESCRIPCION & "'" & ",'" & IMPORTE & "'," & "'" & SITUACION & "'" & "," & _
                        "'" & CONTA & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRemesas_documentos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE REMESAS_DOCUMENTOS SET " & _
                        " REMESA_ID = " & REMESA_ID & "," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " VENCIMIENTO = " & VENCIMIENTO & "," & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " FECHA_VENCIMIENTO = '" & FECHA_VENCIMIENTO & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " IMPORTE = '" & IMPORTE & "'," & _
                        " SITUACION = '" & SITUACION & "'," & _
                        " CONTA = '" & CONTA & "'" & _
                " WHERE ID = " & ID
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
        consulta = "DELETE FROM REMESAS_DOCUMENTOS " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRemesas_documentos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REMESAS_DOCUMENTOS ORDER BY ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,REMESA_ID FROM REMESAS_DOCUMENTOS ORDER BY REMESA_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
