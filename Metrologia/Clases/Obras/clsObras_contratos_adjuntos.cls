VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObras_contratos_adjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOBRAS_CONTRATOS_ADJUNTOS
'***************************************************************
Private ID_ADJUNTO As Long
Private CONTRATO_ID As Long
Private DESCRIPCION As String
Private DOCUMENTO As String
Private FS As String
Private USUARIO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOBRAS_CONTRATOS_ADJUNTOS
'***************************************************************
Public Property Let setID_ADJUNTO(ByVal dato As Long)
        ID_ADJUNTO = dato
End Property
Public Property Let setCONTRATO_ID(ByVal dato As Long)
        CONTRATO_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setDOCUMENTO(ByVal dato As String)
        DOCUMENTO = dato
End Property
Public Property Let setFS(ByVal dato As String)
        FS = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOBRAS_CONTRATOS_ADJUNTOS
'***************************************************************
Public Property Get getID_ADJUNTO() As Long
        getID_ADJUNTO = ID_ADJUNTO
End Property
Public Property Get getCONTRATO_ID() As Long
        getCONTRATO_ID = CONTRATO_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getDOCUMENTO() As String
        getDOCUMENTO = DOCUMENTO
End Property
Public Property Get getFS() As String
        getFS = FS
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOBRAS_CONTRATOS_ADJUNTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OBRAS_CONTRATOS_ADJUNTOS WHERE ID_ADJUNTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ADJUNTO = rs("ID_ADJUNTO")
                CONTRATO_ID = rs("CONTRATO_ID")
                DESCRIPCION = rs("DESCRIPCION")
                DOCUMENTO = rs("DOCUMENTO")
                FS = rs("FS")
                USUARIO_ID = rs("USUARIO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsObras_contratos_adjuntos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_ADJUNTO) FROM OBRAS_CONTRATOS_ADJUNTOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ADJUNTO = 1
        Else
                ID_ADJUNTO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO OBRAS_CONTRATOS_ADJUNTOS " & _
                   "  VALUES (" & _
                        ID_ADJUNTO & "," & CONTRATO_ID & "," & "'" & DESCRIPCION & "'" & "," & _
                        "'" & DOCUMENTO & "'" & ",CURRENT_TIMESTAMP," & USUARIO_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_ADJUNTO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsObras_contratos_adjuntos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OBRAS_CONTRATOS_ADJUNTOS SET " & _
                        " CONTRATO_ID = " & CONTRATO_ID & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " DOCUMENTO = '" & DOCUMENTO & "'," & _
                        " FS = '" & FS & "'," & _
                        " USUARIO_ID = " & USUARIO_ID & "" & _
                " WHERE ID_ADJUNTO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsObras_contratos_adjuntos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OBRAS_CONTRATOS_ADJUNTOS " & _
                "    WHERE ID_ADJUNTO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsObras_contratos_adjuntos)", vbCritical, Err.Description
End Function
Public Function Listado(CONTRATO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_ADJUNTO,A.FS,A.DESCRIPCION,B.USUARIO,A.DOCUMENTO " & _
                   "  FROM OBRAS_CONTRATOS_ADJUNTOS A, USUARIOS B " & _
                   " WHERE A.USUARIO_ID = B.ID_EMPLEADO " & _
                   "   AND A.CONTRATO_ID = " & CONTRATO & _
                   " ORDER BY FS DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_ADJUNTO,CONTRATO_ID FROM OBRAS_CONTRATOS_ADJUNTOS ORDER BY CONTRATO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
