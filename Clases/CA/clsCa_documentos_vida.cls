VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_documentos_vida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_DOCUMENTOS_VIDA
'***************************************************************
Private ID As Long
Private TIPO_ID As Long
Private SUBTIPO_ID As Long
Private IDENTIFICADOR As Long
Private USUARIO As Long
Private fecha As String
Private HORA As String
Private MOTIVO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_DOCUMENTOS_VIDA
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setSUBTIPO_ID(ByVal dato As Long)
        SUBTIPO_ID = dato
End Property
Public Property Let setIDENTIFICADOR(ByVal dato As Long)
        IDENTIFICADOR = dato
End Property
Public Property Let setUSUARIO(ByVal dato As Long)
        USUARIO = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setHORA(ByVal dato As String)
        HORA = dato
End Property
Public Property Let setMOTIVO(ByVal dato As String)
        MOTIVO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_DOCUMENTOS_VIDA
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getSUBTIPO_ID() As Long
        getSUBTIPO_ID = SUBTIPO_ID
End Property
Public Property Get getIDENTIFICADOR() As Long
        getIDENTIFICADOR = IDENTIFICADOR
End Property
Public Property Get getUSUARIO() As Long
        getUSUARIO = USUARIO
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getHORA() As String
        getHORA = HORA
End Property
Public Property Get getMOTIVO() As String
        getMOTIVO = MOTIVO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_DOCUMENTOS_VIDA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_DOCUMENTOS_VIDA WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                TIPO_ID = rs("TIPO_ID")
                SUBTIPO_ID = rs("SUBTIPO_ID")
                IDENTIFICADOR = rs("IDENTIFICADOR")
                USUARIO = rs("USUARIO")
                fecha = rs("FECHA")
                HORA = rs("HORA")
                MOTIVO = rs("MOTIVO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_documentos_vida)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM CA_DOCUMENTOS_VIDA"
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
        consulta = "INSERT INTO CA_DOCUMENTOS_VIDA " & _
                   "  VALUES (" & _
                        ID & "," & TIPO_ID & "," & SUBTIPO_ID & "," & _
                        IDENTIFICADOR & "," & USUARIO & ",CURRENT_DATE,CURRENT_TIME," & _
                        "'" & MOTIVO & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_documentos_vida)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_DOCUMENTOS_VIDA SET " & _
                        " TIPO_ID = " & TIPO_ID & "," & _
                        " SUBTIPO_ID = " & SUBTIPO_ID & "," & _
                        " IDENTIFICADOR = " & IDENTIFICADOR & "," & _
                        " USUARIO = " & USUARIO & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " HORA = '" & HORA & "'," & _
                        " MOTIVO = '" & MOTIVO & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_documentos_vida)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CA_DOCUMENTOS_VIDA " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_documentos_vida)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_DOCUMENTOS_VIDA ORDER BY ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,TIPO_ID FROM CA_DOCUMENTOS_VIDA ORDER BY TIPO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
