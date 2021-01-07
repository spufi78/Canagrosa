VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_responsables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_RESPONSABLES
'***************************************************************
Private ID_RESPONSABLE As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_RESPONSABLES
'***************************************************************
Public Property Let setID_RESPONSABLE(ByVal dato As Long)
        ID_RESPONSABLE = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_RESPONSABLES
'***************************************************************
Public Property Get getID_RESPONSABLE() As Long
        getID_RESPONSABLE = ID_RESPONSABLE
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_RESPONSABLES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_RESPONSABLES WHERE ID_RESPONSABLE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_RESPONSABLE = rs("ID_RESPONSABLE")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_responsables)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_RESPONSABLE) FROM CA_RESPONSABLES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_RESPONSABLE = 1
        Else
                ID_RESPONSABLE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        consulta = "SELECT * FROM CA_RESPONSABLES WHERE NOMBRE = '" & NOMBRE & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            Me.CrearID
            consulta = "INSERT INTO CA_RESPONSABLES " & _
                       "  VALUES (" & _
                            ID_RESPONSABLE & "," & "'" & NOMBRE & "'" & _
                    ")"
            execute_bd consulta
           log ("**********NUEVO RESPONSABLE: " & NOMBRE)
        End If
        Insertar = ID_RESPONSABLE
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_responsables)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_RESPONSABLES SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_RESPONSABLE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_responsables)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        Set rs = datos_bd("select * from ca_documentos where responsable_id = " & ID)
        If rs.RecordCount = 0 Then
            consulta = "DELETE FROM CA_RESPONSABLES " & _
                    "    WHERE ID_RESPONSABLE = " & ID
            execute_bd consulta
            Eliminar = True
        Else
            MsgBox "No se puede eliminar el responsable, tiene documentos asignados.", vbCritical, App.Title
            Eliminar = False
        End If
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_responsables)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_RESPONSABLES ORDER BY ID_RESPONSABLE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_RESPONSABLE,NOMBRE FROM CA_RESPONSABLES ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
