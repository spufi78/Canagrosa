VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComercial"
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
Private ID_COMERCIAL As Long
Private NOMBRE As String
Private APODO As String
Private DIRECCION As String
Private CP As Long
Private PROVINCIA_ID As Long
Private MUNICIPIO_ID As Long
Private CIF As String
Private TELEFONO As String
Private movil As String
Private foto As String
Private OBSERVACIONES As String
Private ESTADO_ID As Integer
Private COMISION As String
Private CODIGO_INTERNO As String
Private ccc As String
Private FAX As String
'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE MUESTRA    *
'*****************************************************
Public Property Let setID_COMERCIAL(ByVal dato As Long)
    ID_COMERCIAL = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setAPODO(ByVal dato As String)
    APODO = Trim(dato)
End Property
Public Property Let setDIRECCION(ByVal dato As String)
    DIRECCION = Trim(dato)
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
    PROVINCIA_ID = dato
End Property
Public Property Let setCP(ByVal dato As Long)
    CP = dato
End Property
Public Property Let setCIF(ByVal dato As String)
    CIF = Trim(dato)
End Property
Public Property Let setTELEFONO(ByVal dato As String)
    TELEFONO = Trim(dato)
End Property
Public Property Let setMOVIL(ByVal dato As String)
    movil = Trim(dato)
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
    MUNICIPIO_ID = dato
End Property
Public Property Let setfoto(ByVal dato As String)
    foto = Trim(dato)
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setESTADO_ID(ByVal dato As Integer)
    ESTADO_ID = Trim(dato)
End Property
Public Property Let setCOMISION(ByVal dato As String)
    COMISION = Trim(dato)
End Property
Public Property Let setCODIGO_INTERNO(ByVal dato As String)
    CODIGO_INTERNO = Trim(dato)
End Property
Public Property Let setCCC(ByVal dato As String)
    ccc = Trim(dato)
End Property
Public Property Let setFAX(ByVal dato As String)
    FAX = Trim(dato)
End Property
Public Property Get getID_COMERCIAL() As Long
    getID_COMERCIAL = ID_COMERCIAL
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getAPODO() As String
    getAPODO = APODO
End Property
Public Property Get getDIRECCION() As String
    getDIRECCION = DIRECCION
End Property
Public Property Get getPROVINCIA_ID() As Long
    getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getCP() As Long
    getCP = CP
End Property
Public Property Get getCIF() As String
    getCIF = CIF
End Property
Public Property Get getTELEFONO() As String
    getTELEFONO = TELEFONO
End Property
Public Property Get getMOVIL() As String
    getMOVIL = movil
End Property
Public Property Get getMUNICIPIO_ID() As Long
    getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getfoto() As String
    getfoto = foto
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getESTADO_ID() As String
    getESTADO_ID = ESTADO_ID
End Property
Public Property Get getCOMISION() As String
    getCOMISION = COMISION
End Property
Public Property Get getCODIGO_INTERNO() As String
    getCODIGO_INTERNO = CODIGO_INTERNO
End Property
Public Property Get getCCC() As String
    getCCC = ccc
End Property
Public Property Get getFAX() As String
    getFAX = FAX
End Property
Public Function Cargar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM COMERCIALES WHERE ID_COMERCIAL=" & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Cargar = False
        Else
            ID_COMERCIAL = rs.Fields("ID_COMERCIAL")
            NOMBRE = rs.Fields("NOMBRE")
            APODO = rs.Fields("APODO")
            DIRECCION = rs.Fields("DIRECCION")
            PROVINCIA_ID = rs.Fields("PROVINCIA_ID")
            If Not IsNull(rs.Fields("CP")) Then
                CP = rs.Fields("CP")
            Else
                CP = 0
            End If
            CIF = rs.Fields("CIF")
            TELEFONO = rs.Fields("TELEFONO")
            movil = rs.Fields("MOVIL")
            MUNICIPIO_ID = rs.Fields("MUNICIPIO_ID")
            foto = rs.Fields("FOTO")
            OBSERVACIONES = rs.Fields("OBSERVACIONES")
            ESTADO_ID = rs.Fields("ESTADO_ID")
            COMISION = rs("COMISION")
            CODIGO_INTERNO = rs.Fields("CODIGO_INTERNO")
            ccc = rs.Fields("CCC")
            FAX = rs.Fields("FAX")
            Cargar = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Cargar = False
        MsgBox "Error al cargar los datos del operario (Cargar)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
    Dim consulta As String
    On Error GoTo fallo:
    If ID_COMERCIAL = 0 Then
        Me.CrearID_COMERCIAL
    End If
    consulta = "INSERT INTO COMERCIALES" & _
               " VALUES(" & _
               ID_COMERCIAL & ",'" & NOMBRE & "','" & APODO & "','" & DIRECCION & "'," & _
               CP & "," & PROVINCIA_ID & "," & MUNICIPIO_ID & ",'" & _
               CIF & "','" & TELEFONO & "','" & movil & "','" & _
               foto & "','" & OBSERVACIONES & "'," & ESTADO_ID & ",'" & COMISION & "','" & CODIGO_INTERNO & "','" & _
               ccc & "','" & FAX & "')"
    execute_bd consulta
    Insertar = ID_COMERCIAL
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al guardar el operario en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Sub CrearID_COMERCIAL()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_COMERCIAL) FROM COMERCIALES"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_COMERCIAL = 1
    Else
        ID_COMERCIAL = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Modificar(ID As Integer) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE COMERCIALES SET" & _
               "  NOMBRE = '" & NOMBRE & "'," & _
               "  APODO = '" & APODO & "'," & _
               "  DIRECCION = '" & DIRECCION & "'," & _
               "  CP = " & CP & "," & _
               "  PROVINCIA_ID = " & PROVINCIA_ID & "," & _
               "  MUNICIPIO_ID = " & MUNICIPIO_ID & "," & _
               "  CIF = '" & CIF & "'," & _
               "  TELEFONO = '" & TELEFONO & "'," & _
               "  MOVIL = '" & movil & "'," & _
               "  FOTO = '" & foto & "'," & _
               "  OBSERVACIONES = '" & OBSERVACIONES & "'," & _
               "  CODIGO_INTERNO = '" & CODIGO_INTERNO & "'," & _
               "  COMISION = '" & COMISION & "'," & _
               "  CCC = '" & ccc & "'," & _
               "  FAX = '" & FAX & "'," & _
               "  ESTADO_ID = " & ESTADO_ID & _
               " WHERE ID_COMERCIAL = " & ID
               execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el operario en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Eliminar(ID As Integer) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "DELETE FROM COMERCIALES " & _
               " WHERE ID_COMERCIAL = " & ID
               execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el operario de la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM COMERCIALES order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Estado(ESTADO As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM COMERCIALES where ESTADO_ID = " & ESTADO & " order by nombre"
    Set Listado_Estado = datos_bd(consulta)
End Function
Public Function Listado_Apodos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT apodo,ID_COMERCIAL,ESTADO_ID FROM COMERCIALES order by apodo"
    Set Listado_Apodos = datos_bd(consulta)
End Function
Public Function Listado_Apodos_Activos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT apodo,ID_COMERCIAL FROM COMERCIALES WHERE ESTADO_ID = 0 order by apodo"
    Set Listado_Apodos_Activos = datos_bd(consulta)
End Function
Public Function Listado_Nominas() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM COMERCIALES where ESTADO_ID <> 2 order by codigo_interno asc"
    Set Listado_Nominas = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_COMERCIAL, NOMBRE FROM COMERCIALES " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
