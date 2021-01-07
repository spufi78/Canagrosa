VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_estados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_ESTADOS
'***************************************************************
Private ID_ESTADO As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_ESTADOS
'***************************************************************
Public Property Let setID_ESTADO(ByVal dato As Long)
        ID_ESTADO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_ESTADOS
'***************************************************************
Public Property Get getID_ESTADO() As Long
        getID_ESTADO = ID_ESTADO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_ESTADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_ESTADOS WHERE ID_ESTADO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ESTADO = rs("ID_ESTADO")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_estados)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_ESTADO) FROM CA_ESTADOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ESTADO = 1
        Else
                ID_ESTADO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        consulta = "SELECT * FROM CA_ESTADOS WHERE NOMBRE = '" & NOMBRE & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            Me.CrearID
            consulta = "INSERT INTO CA_ESTADOS " & _
                       "  VALUES (" & _
                            ID_ESTADO & "," & "'" & NOMBRE & "'" & _
                    ")"
            execute_bd consulta
        End If
        Insertar = ID_ESTADO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_estados)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_ESTADOS SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_ESTADO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_estados)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        Set rs = datos_bd("select * from ca_documentos where estado_id = " & ID)
        If rs.RecordCount = 0 Then
            consulta = "DELETE FROM CA_ESTADOS " & _
                    "    WHERE ID_ESTADO = " & ID
            execute_bd consulta
            Eliminar = True
        Else
            MsgBox "No se puede eliminar el estado, tiene documentos asignados.", vbCritical, App.Title
            Eliminar = False
        End If
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_Estados)", vbCritical, Err.Description

End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_ESTADOS ORDER BY ID_ESTADO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_ESTADO,NOMBRE FROM CA_ESTADOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
