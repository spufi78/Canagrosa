VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_familias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_FAMILIAS
'***************************************************************
Private ID_FAMILIA As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_FAMILIAS
'***************************************************************
Public Property Let setID_FAMILIA(ByVal dato As Long)
        ID_FAMILIA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_FAMILIAS
'***************************************************************
Public Property Get getID_FAMILIA() As Long
        getID_FAMILIA = ID_FAMILIA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_FAMILIAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_FAMILIAS WHERE ID_FAMILIA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FAMILIA = rs("ID_FAMILIA")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_familias)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_FAMILIA) FROM CA_FAMILIAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FAMILIA = 1
        Else
                ID_FAMILIA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        consulta = "SELECT * FROM CA_FAMILIAS WHERE NOMBRE = '" & NOMBRE & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            Me.CrearID
            consulta = "INSERT INTO CA_FAMILIAS " & _
                       "  VALUES (" & _
                            ID_FAMILIA & "," & "'" & NOMBRE & "'" & _
                    ")"
            execute_bd consulta
            log ("**********NUEVA FAMILIA: " & NOMBRE)

        End If
        Insertar = ID_FAMILIA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_familias)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_FAMILIAS SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_FAMILIA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_familias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        Set rs = datos_bd("select * from ca_documentos where familia_id = " & ID)
        If rs.RecordCount = 0 Then
            consulta = "DELETE FROM CA_FAMILIAS " & _
                    "    WHERE ID_FAMILIA = " & ID
            execute_bd consulta
            Eliminar = True
        Else
            MsgBox "No se puede eliminar la familia, tiene documentos asignados.", vbCritical, App.Title
            Eliminar = False
        End If
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_familias)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_FAMILIAS ORDER BY ID_FAMILIA"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_FAMILIA,NOMBRE FROM CA_FAMILIAS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
