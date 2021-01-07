VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSREMESAS
'***************************************************************
Private ID_REMESA As Long
Private fecha As String
Private BANCO_ID As Long
Private USUARIO_ID As Long
Private TIMESTAMP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSREMESAS
'***************************************************************
Public Property Let setID_REMESA(ByVal dato As Long)
        ID_REMESA = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setBANCO_ID(ByVal dato As Long)
        BANCO_ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTIMESTAMP(ByVal dato As String)
        TIMESTAMP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSREMESAS
'***************************************************************
Public Property Get getID_REMESA() As Long
        getID_REMESA = ID_REMESA
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getBANCO_ID() As Long
        getBANCO_ID = BANCO_ID
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTIMESTAMP() As String
        getTIMESTAMP = TIMESTAMP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSREMESAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REMESAS WHERE ID_REMESA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_REMESA = rs("ID_REMESA")
                fecha = rs("FECHA")
                BANCO_ID = rs("BANCO_ID")
                USUARIO_ID = rs("USUARIO_ID")
                TIMESTAMP = rs("TIMESTAMP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRemesas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_REMESA) FROM REMESAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_REMESA = 1
        Else
                ID_REMESA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_REMESA = 0 Then
            Me.CrearID
        End If
        consulta = "INSERT INTO REMESAS " & _
                   "  VALUES (" & _
                        ID_REMESA & "," & "'" & fecha & "'" & "," & BANCO_ID & "," & _
                        USUARIO_ID & "," & " current_timestamp " & _
                ")"
        execute_bd consulta
        Insertar = ID_REMESA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRemesas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE REMESAS SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " BANCO_ID = " & BANCO_ID & "," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TIMESTAMP = current_timestamp " & _
                " WHERE ID_REMESA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRemesas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT * FROM REMESAS_DOCUMENTOS WHERE REMESA_ID = " & ID & " AND SITUACION = 'D'"
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar, la remesa tiene efectos descontados.", vbCritical, App.Title
            Exit Function
        End If
        execute_bd "DELETE FROM REMESAS WHERE ID_REMESA = " & ID
        execute_bd "DELETE FROM REMESAS_DOCUMENTOS WHERE REMESA_ID = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRemesas)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REMESAS ORDER BY ID_REMESA"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_REMESA,FECHA FROM REMESAS ORDER BY FECHA"
        Set Listado_Combo = datos_bd(consulta)
End Function
