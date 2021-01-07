VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocs_pago_envios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDOCS_PAGO_ENVIOS
'***************************************************************
Private DOC_ID As Long
Private ORDEN As Long
Private fecha As String
Private DESCRIPCION As String
Private USUARIO_ID As Long
Private TS As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDOCS_PAGO_ENVIOS
'***************************************************************
Public Property Let setDOC_ID(ByVal dato As Long)
        DOC_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDOCS_PAGO_ENVIOS
'***************************************************************
Public Property Get getDOC_ID() As Long
        getDOC_ID = DOC_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDOCS_PAGO_ENVIOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCS_PAGO_ENVIOS WHERE DOC_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOC_ID = rs("DOC_ID")
                ORDEN = rs("ORDEN")
                fecha = rs("FECHA")
                DESCRIPCION = rs("DESCRIPCION")
                USUARIO_ID = rs("USUARIO_ID")
                TS = rs("TS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDocs_pago_envios)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO DOCS_PAGO_ENVIOS " & _
                   "  VALUES (" & _
                        DOC_ID & "," & ORDEN & "," & "'" & fecha & "'" & "," & _
                        "'" & DESCRIPCION & "'," & USUARIO.getID_EMPLEADO & ",CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Dim od As New clsDocs_pago
        od.Enviar DOC_ID, 1
        Set od = Nothing
        Insertar = DOC_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDocs_pago_envios)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCS_PAGO_ENVIOS SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE DOC_ID = " & ID & " AND ORDEN = " & ORDEN
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDocs_pago_envios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCS_PAGO_ENVIOS " & _
                "    WHERE DOC_ID = " & ID & " AND ORDEN = " & ORDEN
        execute_bd consulta
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("SELECT * FROM DOCS_PAGO_ENVIOS WHERE DOC_ID = " & ID)
        If rs.RecordCount = 0 Then
            Dim od As New clsDocs_pago
            od.Enviar ID, 0
            Set od = Nothing
        End If
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDocs_pago_envios)", vbCritical, Err.Description
End Function
Public Function Listado(PK As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ORDEN, A.FECHA,A.DESCRIPCION,U.USUARIO,A.TS " & _
                   "  FROM DOCS_PAGO_ENVIOS A " & _
                   "  LEFT JOIN usuarios U ON A.USUARIO_ID = U.ID_EMPLEADO " & _
                   " WHERE A.DOC_ID = " & PK & _
                   " ORDER BY A.FECHA,A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
