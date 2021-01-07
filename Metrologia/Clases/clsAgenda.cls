VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLIENTE      *
'*****************************************************
Private ID_AGENDA As Long
Private NOMBRE As String
Private TELEFONO As String
Private movil As String
Private FAX As String
Private CLIENTE_ID As Long
'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLIENTES   *
'*****************************************************
Public Property Let setID_AGENDA(ByVal dato As Long)
    ID_AGENDA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setTELEFONO(ByVal dato As String)
    TELEFONO = Trim(dato)
End Property
Public Property Let setMOVIL(ByVal dato As String)
    movil = Trim(dato)
End Property
Public Property Let setFAX(ByVal dato As String)
    FAX = Trim(dato)
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
    CLIENTE_ID = dato
End Property
'*****************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLIENTES   *
'*****************************************************
Public Property Get getID_AGENDA() As Long
    getID_AGENDA = ID_AGENDA
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getTELEFONO() As String
    getTELEFONO = TELEFONO
End Property
Public Property Get getMOVIL() As String
    getMOVIL = movil
End Property
Public Property Get getFAX() As String
    getFAX = FAX
End Property
Public Property Get getCLIENTE_ID() As Long
    getCLIENTE_ID = CLIENTE_ID
End Property
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM AGENDA WHERE ID_AGENDA=" & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            ID_AGENDA = rs.Fields("ID_AGENDA")
            NOMBRE = rs.Fields("NOMBRE") & ""
            TELEFONO = rs.Fields("TELEFONO") & ""
            movil = rs.Fields("MOVIL") & ""
            FAX = rs.Fields("FAX")
            CLIENTE_ID = rs.Fields("CLIENTE_ID")
            Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos de la agenda (Carga)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
    Dim consulta As String
    On Error GoTo fallo:
    Me.CrearID
    consulta = "INSERT INTO AGENDA" & _
               " VALUES(" & _
               ID_AGENDA & ",'" & UCase(NOMBRE) & "','" & _
               TELEFONO & "','" & movil & "','" & FAX & "'," & CLIENTE_ID & ")"
    execute_bd consulta
    Insertar = ID_AGENDA
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al guardar en la agenda : " & Err.Description, vbCritical, Err.Number
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_AGENDA) FROM AGENDA"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_AGENDA = 1
    Else
        ID_AGENDA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Modificar() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE AGENDA SET" & _
               " NOMBRE = '" & UCase(NOMBRE) & "'," & _
               " TELEFONO = '" & TELEFONO & "'," & _
               " MOVIL = '" & movil & "'," & _
               " FAX = '" & FAX & "'" & _
               " WHERE " & _
               " ID_AGENDA = " & ID_AGENDA
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar en la agenda : " & Err.Description, vbCritical, Err.Number
End Function
Public Function modificar_cliente(cliente As Long) As Boolean
    Dim consulta As String
    Dim rs As ADODB.Recordset
    On Error GoTo fallo:
    consulta = "SELECT TELEFONO FROM AGENDA WHERE CLIENTE_ID = " & cliente
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        consulta = "UPDATE AGENDA SET" & _
                   " TELEFONO = '" & TELEFONO & "'," & _
                   " MOVIL = '" & movil & "'," & _
                   " FAX = '" & FAX & "'" & _
                   " WHERE CLIENTE_ID = " & cliente
        execute_bd consulta
    Else
        Me.Insertar
    End If
    modificar_cliente = True
    Exit Function
fallo:
    modificar_cliente = False
    MsgBox "Error al modificar en la agenda : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Eliminar() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "DELETE FROM AGENDA " & _
               " WHERE " & _
               " ID_AGENDA = " & ID_AGENDA
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar de la agenda : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM AGENDA order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_letra(letra As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT NOMBRE,TELEFONO,MOVIL,FAX,ID_AGENDA FROM AGENDA WHERE NOMBRE LIKE '%" & letra & "%' ORDER BY NOMBRE"
    Set Listado_por_letra = datos_bd(consulta)
End Function

