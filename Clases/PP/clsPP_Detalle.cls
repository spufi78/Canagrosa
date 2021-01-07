VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPP_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsPP_Detalle
'***************************************************************
Private ID As Long
Private PP_ID As Long
Private REFERENCIA As String
Private FAMILIA_ID As String
Private DESCRIPCION As String
Private unidades As String
Private DESCUENTO As Long
Private PRECIO As String
Private IMPORTE As String
Private CUSERID As Long
Private FTIMESTP As String

Private R_UNIDADES As String
Private R_USUARIO_ID As Long
Private R_TS As String


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsPP_Detalle
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setPP_ID(ByVal dato As Long)
        PP_ID = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Long)
        FAMILIA_ID = dato
End Property
Public Property Let setREFERENCIA(ByVal dato As String)
        REFERENCIA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setUNIDADES(ByVal dato As String)
        unidades = dato
End Property
Public Property Let setDESCUENTO(ByVal dato As Long)
        DESCUENTO = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
Public Property Let setCUSERID(ByVal dato As Long)
        CUSERID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property

Public Property Let setR_UNIDADES(ByVal dato As String)
        R_UNIDADES = dato
End Property
Public Property Let setR_USUARIO_ID(ByVal dato As Long)
        R_USUARIO_ID = dato
End Property
Public Property Let setR_TS(ByVal dato As String)
        R_TS = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsPP_Detalle
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getPP_ID() As Long
        getPP_ID = PP_ID
End Property
Public Property Get getREFERENCIA() As String
        getREFERENCIA = REFERENCIA
End Property
Public Property Get getFAMILIA_ID() As Long
        getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getUNIDADES() As String
        getUNIDADES = unidades
End Property
Public Property Get getDESCUENTO() As String
        getDESCUENTO = DESCUENTO
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
Public Property Get getCUSERID() As Long
        getCUSERID = CUSERID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property

Public Property Get getR_UNIDADES() As String
        getR_UNIDADES = R_UNIDADES
End Property
Public Property Get getR_USUARIO_ID() As Long
        getR_USUARIO_ID = R_USUARIO_ID
End Property
Public Property Get getR_TS() As String
        getR_TS = R_TS
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsPP_Detalle
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PP_DETALLE WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                PP_ID = rs("PP_ID")
                REFERENCIA = rs("REFERENCIA")
                FAMILIA_ID = rs("FAMILIA_ID")
                DESCRIPCION = rs("DESCRIPCION")
                unidades = rs("UNIDADES")
                DESCUENTO = rs("DESCUENTO")
                PRECIO = rs("PRECIO")
                IMPORTE = rs("IMPORTE")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                
                R_UNIDADES = rs("R_UNIDADES")
                R_USUARIO_ID = rs("R_USUARIO_ID")
                R_TS = rs("R_TS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPP_Detalle)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        CrearID
        consulta = "INSERT INTO PP_DETALLE " & _
                   "  VALUES (" & ID & "," & _
                        PP_ID & ",'" & REFERENCIA & "'," & FAMILIA_ID & ",'" & DESCRIPCION & "'" & ",'" & unidades & "'," & _
                        DESCUENTO & ",'" & PRECIO & "'" & ",'" & IMPORTE & "'," & USUARIO.getID_EMPLEADO & ", CURRENT_TIMESTAMP,0,0,NULL" & _
                ") ON DUPLICATE KEY UPDATE DESCRIPCION = '" & DESCRIPCION & "', FAMILIA_ID = " & FAMILIA_ID & ", UNIDADES = '" & unidades & "', PRECIO = '" & PRECIO & "', CUSERID = " & USUARIO.getID_EMPLEADO & ", FTIMESTP = CURRENT_TIMESTAMP"
        execute_bd consulta
        Insertar = PP_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPP_Detalle)", vbCritical, Err.Description
End Function
Public Function recepcionar(ID As Long, unidades As Double) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PP_DETALLE " & _
                   "   SET R_UNIDADES = '" & moneda_bd(CStr(unidades)) & "'" & _
                   "     , R_USUARIO_ID = " & USUARIO.getID_EMPLEADO & _
                   "     , R_TS = CURRENT_TIMESTAMP " & _
                   " WHERE ID = " & ID
        execute_bd consulta
        recepcionar = True
        Exit Function
fallo:
        recepcionar = False
        MsgBox "Error al recepcionar (clsPP_Detalle)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PP_DETALLE " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPP_Detalle)", vbCritical, Err.Description
End Function
Public Function Eliminar_Pedido(PK As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PP_DETALLE " & _
                "    WHERE PP_ID = " & PK
        execute_bd consulta
        Eliminar_Pedido = True
        Exit Function
fallo:
        Eliminar_Pedido = False
        MsgBox "Error al Eliminar_Pedido (clsPP_Detalle)", vbCritical, Err.Description
End Function
Private Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    On Error GoTo fallo
    ' Recuperamos el maximo secuencial de muestra
    consulta = "SELECT MAX(ID) FROM PP_DETALLE "

    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID = 1
    Else
        ID = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error en CrearId : " & Err.Description, vbCritical, App.Title
End Sub

Public Function Listado(ID_PP As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT A.*,B.NOMBRE AS FAMILIA,C.NOMBRE AS USUARIO " & _
               "  FROM PP_DETALLE A " & _
               "  LEFT JOIN FAMILIAS B ON A.FAMILIA_ID = B.ID_FAMILIA " & _
               "  LEFT JOIN usuarios C ON A.R_USUARIO_ID = C.ID_EMPLEADO " & _
               " WHERE PP_ID =  " & ID_PP & _
               " ORDER BY ID "
    Set Listado = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function


Public Function ListadoDescripciones() As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT DISTINCT DESCRIPCION,DESCUENTO,PRECIO " & _
               "  FROM PP_DETALLE " & _
               " ORDER BY DESCRIPCION "
    Set ListadoDescripciones = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function


