VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_tipos_ensayos_contratas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_TIPOS_ENSAYOS_CONTRATAS
'***************************************************************
Private ID_TIPO_ENSAYO_CONTRATA As Long
Private TIPO_ENSAYO_ID As Long
Private CONTRATA_ID As Long
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_TIPOS_ENSAYOS_CONTRATAS
'***************************************************************
Public Property Let setID_TIPO_ENSAYO_CONTRATA(ByVal dato As Long)
        ID_TIPO_ENSAYO_CONTRATA = dato
End Property
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setCONTRATA_ID(ByVal dato As Long)
        CONTRATA_ID = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_TIPOS_ENSAYOS_CONTRATAS
'***************************************************************
Public Property Get getID_TIPO_ENSAYO_CONTRATA() As Long
        getID_TIPO_ENSAYO_CONTRATA = ID_TIPO_ENSAYO_CONTRATA
End Property
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getCONTRATA_ID() As Long
        getCONTRATA_ID = CONTRATA_ID
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_TIPOS_ENSAYOS_CONTRATAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_TIPOS_ENSAYOS_CONTRATAS WHERE ID_TIPO_ENSAYO_CONTRATA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TIPO_ENSAYO_CONTRATA = rs("ID_TIPO_ENSAYO_CONTRATA")
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                CONTRATA_ID = rs("CONTRATA_ID")
                PRECIO = rs("PRECIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_tipos_ensayos_contratas)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        CrearId
        consulta = "INSERT INTO CE_TIPOS_ENSAYOS_CONTRATAS " & _
                   "  VALUES (" & _
                        ID_TIPO_ENSAYO_CONTRATA & "," & TIPO_ENSAYO_ID & "," & CONTRATA_ID & "," & _
                        "'" & PRECIO & "')" & _
                         " ON DUPLICATE KEY UPDATE CONTRATA_ID= " & CONTRATA_ID
        execute_bd consulta
        Insertar = ID_TIPO_ENSAYO_CONTRATA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_tipos_ensayos_contratas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_TIPOS_ENSAYOS_CONTRATAS SET " & _
                        " TIPO_ENSAYO_ID = " & TIPO_ENSAYO_ID & "," & _
                        " CONTRATA_ID = " & CONTRATA_ID & "," & _
                        " PRECIO = '" & PRECIO & "'" & _
                " WHERE ID_TIPO_ENSAYO_CONTRATA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_tipos_ensayos_contratas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_TIPOS_ENSAYOS_CONTRATAS " & _
                "    WHERE ID_TIPO_ENSAYO_CONTRATA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_tipos_ensayos_contratas)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_TIPOS_ENSAYOS_CONTRATAS ORDER BY ID_TIPO_ENSAYO_CONTRATA"
        Set Listado = datos_bd(consulta)
End Function

Private Sub CrearId()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    On Error GoTo fallo
    ' Recuperamos el maximo
    consulta = "SELECT MAX(ID_TIPO_ENSAYO_CONTRATA) FROM ce_tipos_ensayos_contratas "
              
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_ENSAYO_CONTRATA = 1
    Else
        ID_TIPO_ENSAYO_CONTRATA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error en CrearID_TIPO_ENSAYO_CONTRATA : " & Err.Description, vbCritical, App.Title
End Sub

