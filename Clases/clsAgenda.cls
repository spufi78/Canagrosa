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
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLIENTE      *
'*****************************************************
Private ID_AGENDA As Long
Private nombre As String
Private telefono As String
Private movil As String
Private FAX As String
'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLIENTES   *
'*****************************************************
Public Property Let setID_AGENDA(ByVal dato As Long)
    ID_AGENDA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setTELEFONO(ByVal dato As String)
    telefono = Trim(dato)
End Property
Public Property Let setMOVIL(ByVal dato As String)
    movil = Trim(dato)
End Property
Public Property Let setFAX(ByVal dato As String)
    FAX = Trim(dato)
End Property
'*****************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLIENTES   *
'*****************************************************
Public Property Get getID_AGENDA() As Long
    getID_AGENDA = ID_AGENDA
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getTELEFONO() As String
    getTELEFONO = telefono
End Property
Public Property Get getMOVIL() As String
    getMOVIL = movil
End Property
Public Property Get getFAX() As String
    getFAX = FAX
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
            nombre = rs.Fields("NOMBRE") & ""
            telefono = rs.Fields("TELEFONO") & ""
            movil = rs.Fields("MOVIL") & ""
            FAX = rs.Fields("FAX")
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
               ID_AGENDA & ",'" & nombre & "','" & _
               telefono & "','" & movil & "','" & FAX & "')"
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
               " NOMBRE = '" & nombre & "'," & _
               " TELEFONO = '" & telefono & "'," & _
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
Public Function Listado_por_letra(nombre As String) As ADODB.Recordset
    Dim consulta As String
    Dim filtro As String
    If nombre <> "" Then
        filtro = " AND NOMBRE LIKE '%" & nombre & "%'"
    End If
    consulta = "SELECT NOMBRE,TELEFONO,MOVIL,FAX,ID_AGENDA " & _
               "  FROM AGENDA " & _
               " WHERE 1 = 1 " & _
               filtro & _
               " ORDER BY NOMBRE"
    Set Listado_por_letra = datos_bd(consulta)
End Function
Public Function Listado_Filtro(filtro As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT NOMBRE,TELEFONO,MOVIL,FAX,ID_AGENDA FROM AGENDA WHERE NOMBRE LIKE '%" & filtro & "%' ORDER BY NOMBRE"
    Set Listado_Filtro = datos_bd(consulta)
End Function

