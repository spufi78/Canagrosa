VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForma_pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***********************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE FORMA_PAGO           *
'***********************************************************
Private ID_FORMA_PAGO As Integer
Private NOMBRE As String
Private VENCIMIENTOS As Integer
Private DIAS As Integer
Private DIA_COBRO As Integer
Private APLAZAMIENTO As Integer
Private RECIBOS As Integer
Private TIPO_FORMA As String
Public Property Let setID_FORMA_PAGO(ByVal dato As Integer)
    ID_FORMA_PAGO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setVENCIMIENTOS(ByVal dato As Integer)
    VENCIMIENTOS = dato
End Property
Public Property Let setDIAS(ByVal dato As Integer)
    DIAS = dato
End Property
Public Property Let setDIA_COBRO(ByVal dato As Integer)
    DIA_COBRO = dato
End Property
Public Property Let setAPLAZAMIENTO(ByVal dato As Integer)
    APLAZAMIENTO = dato
End Property
Public Property Let setRECIBOS(ByVal dato As Integer)
    RECIBOS = dato
End Property
Public Property Let setTIPO_FORMA(ByVal dato As String)
    TIPO_FORMA = Trim(dato)
End Property

Public Property Get getID_FORMA_PAGO() As Integer
    getID_FORMA_PAGO = ID_FORMA_PAGO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getVENCIMIENTOS() As Integer
    getVENCIMIENTOS = VENCIMIENTOS
End Property
Public Property Get getDIAS() As Integer
    getDIAS = DIAS
End Property
Public Property Get getDIA_COBRO() As Integer
    getDIA_COBRO = DIA_COBRO
End Property
Public Property Get getAPLAZAMIENTO() As Integer
    getAPLAZAMIENTO = APLAZAMIENTO
End Property
Public Property Get getRECIBOS() As Integer
    getRECIBOS = RECIBOS
End Property
Public Property Get getTIPO_FORMA() As String
    getTIPO_FORMA = TIPO_FORMA
End Property

Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FORMA_PAGO order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_ID(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FORMA_PAGO WHERE ID_FORMA_PAGO = " & ID
    Set Listado_ID = datos_bd(consulta)
End Function
Public Function Cargar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FORMA_PAGO WHERE ID_FORMA_PAGO=" & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID_FORMA_PAGO = rs("ID_FORMA_PAGO")
        NOMBRE = rs("NOMBRE")
        VENCIMIENTOS = rs("VENCIMIENTOS")
        DIAS = rs("DIAS")
        DIA_COBRO = rs("DIA_COBRO")
        APLAZAMIENTO = rs("APLAZAMIENTO")
        RECIBOS = rs("RECIBOS")
        TIPO_FORMA = rs("TIPO_FORMA")
        Cargar = True
    Else
        Cargar = False
    End If
    Exit Function
fallo:
    Cargar = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_FORMA_PAGO) FROM FORMA_PAGO"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_FORMA_PAGO = 1
    Else
        ID_FORMA_PAGO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FORMA_PAGO FROM FORMA_PAGO WHERE NOMBRE = '" & Me.getNOMBRE & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        duplicado = True
    End If
    Set rs = Nothing
End Function

Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    If Me.duplicado = False Then
        Me.CrearID
        consulta = "Insert into FORMA_PAGO " & _
                   " values(" & _
                   ID_FORMA_PAGO & ",'" & NOMBRE & "'," & _
                   VENCIMIENTOS & "," & DIAS & "," & DIA_COBRO & "," & _
                   APLAZAMIENTO & "," & RECIBOS & ",'" & TIPO_FORMA & "')"
        execute_bd consulta
        Insertar = ID_FORMA_PAGO
    Else
        MsgBox "Ya existe una forma de pago con esa descripción.", vbExclamation, App.Title
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el COLOR (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    ' Validar clientes
    consulta = "SELECT * FROM CLIENTES WHERE FORMA_PAGO = " & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        MsgBox "Existen clientes asociados a esta forma de pago. No se puede eliminar.", vbInformation, App.Title
        Exit Function
    End If
    ' Validar clientes
    consulta = "SELECT * FROM DOCUMENTOS WHERE FORMA_PAGO_ID = " & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        MsgBox "Existen documentos asociados a esta forma de pago. No se puede eliminar.", vbInformation, App.Title
        Exit Function
    End If
    ' Eliminar
    consulta = "DELETE FROM FORMA_PAGO WHERE ID_FORMA_PAGO=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function

Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FORMA_PAGO, NOMBRE FROM FORMA_PAGO " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMA_PAGO SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " VENCIMIENTOS = " & VENCIMIENTOS & "," & _
                        " DIA_COBRO = " & DIA_COBRO & "," & _
                        " APLAZAMIENTO = " & APLAZAMIENTO & "," & _
                        " TIPO_FORMA = '" & TIPO_FORMA & "'," & _
                        " RECIBOS = " & RECIBOS & "," & _
                        " DIAS = " & DIAS & _
                " WHERE ID_FORMA_PAGO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (ClsFORMA_PAGO)", vbCritical, Err.Description
End Function
