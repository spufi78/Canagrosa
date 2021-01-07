VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_subfamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_SUBFAMILIAS
'***************************************************************
Private ID_SUBFAMILIA As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_SUBFAMILIAS
'***************************************************************
Public Property Let setID_SUBFAMILIA(ByVal dato As Long)
        ID_SUBFAMILIA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_SUBFAMILIAS
'***************************************************************
Public Property Get getID_SUBFAMILIA() As Long
        getID_SUBFAMILIA = ID_SUBFAMILIA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_SUBFAMILIAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_SUBFAMILIAS WHERE ID_SUBFAMILIA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SUBFAMILIA = rs("ID_SUBFAMILIA")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_subfamilias)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_SUBFAMILIA) FROM CA_SUBFAMILIAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_SUBFAMILIA = 1
        Else
                ID_SUBFAMILIA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        consulta = "SELECT * FROM CA_SUBFAMILIAS WHERE NOMBRE = '" & NOMBRE & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            Me.CrearID
            consulta = "INSERT INTO CA_SUBFAMILIAS " & _
                       "  VALUES (" & _
                            ID_SUBFAMILIA & "," & "'" & NOMBRE & "'" & _
                    ")"
            execute_bd consulta
            log ("**********NUEVA SUBFAMILIA: " & NOMBRE)

        End If
        Insertar = ID_SUBFAMILIA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_subfamilias)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_SUBFAMILIAS SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_SUBFAMILIA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_subfamilias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        Set rs = datos_bd("select * from ca_documentos where subfamilia_id = " & ID)
        If rs.RecordCount = 0 Then
            consulta = "DELETE FROM CA_SUBFAMILIAS " & _
                    "    WHERE ID_SUBFAMILIA = " & ID
            execute_bd consulta
            Eliminar = True
        Else
            MsgBox "No se puede eliminar la subfamilia, tiene documentos asignados.", vbCritical, App.Title
            Eliminar = False
        End If
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_subfamilias)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_SUBFAMILIAS ORDER BY ID_SUBFAMILIA"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_SUBFAMILIA,NOMBRE FROM CA_SUBFAMILIAS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
