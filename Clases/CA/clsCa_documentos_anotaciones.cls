VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_documentos_anotaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_DOCUMENTOS_ANOTACIONES
'***************************************************************
Private DOCUMENTO_ID As Long
Private ORDEN As Long
Private fecha As String
Private USUARIO_ID As Long
Private texto As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_DOCUMENTOS_ANOTACIONES
'***************************************************************
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTEXTO(ByVal dato As String)
        texto = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_DOCUMENTOS_ANOTACIONES
'***************************************************************
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTEXTO() As String
        getTEXTO = texto
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_DOCUMENTOS_ANOTACIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_DOCUMENTOS_ANOTACIONES WHERE DOCUMENTO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                ORDEN = rs("ORDEN")
                fecha = rs("FECHA")
                USUARIO_ID = rs("USUARIO_ID")
                texto = rs("TEXTO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_documentos_anotaciones)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CA_DOCUMENTOS_ANOTACIONES " & _
                   "  VALUES (" & _
                        DOCUMENTO_ID & "," & ORDEN & "," & "'" & fecha & "'" & "," & _
                        USUARIO_ID & "," & "'" & texto & "'" & _
                ")"
        execute_bd consulta
        Insertar = DOCUMENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_documentos_anotaciones)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_DOCUMENTOS_ANOTACIONES SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TEXTO = '" & texto & "'" & _
                " WHERE DOCUMENTO_ID = " & ID & " AND ORDEN = " & ORDEN
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_documentos_anotaciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CA_DOCUMENTOS_ANOTACIONES " & _
                "    WHERE DOCUMENTO_ID = " & ID & " AND ORDEN = " & ORDEN
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_documentos_anotaciones)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.FECHA, B.USUARIO, A.TEXTO, A.ORDEN " & _
                   "  FROM CA_DOCUMENTOS_ANOTACIONES A, usuarios B " & _
                   " WHERE A.USUARIO_ID = B.ID_EMPLEADO " & _
                   "   AND DOCUMENTO_ID = " & ID & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
'M1380-I
Public Function valorMaximoOrden(ID As Long)
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ORDEN) FROM CA_DOCUMENTOS_ANOTACIONES WHERE DOCUMENTO_ID = " & ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Then
            ORDEN = 0
        Else
            ORDEN = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
'M1380-F
Public Function anotaciones(ID As Long) As Integer
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        Dim cantidad As Integer
        cantidad = 0
        consulta = "SELECT COUNT(*) FROM CA_DOCUMENTOS_ANOTACIONES WHERE DOCUMENTO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount <> 0 Then
            cantidad = rs(0)
        End If
        Set rs = Nothing
        anotaciones = cantidad
        Exit Function
fallo:
        anotaciones = 0
        MsgBox "Error al cargar las anotaciones(clsCa_documentos_anotaciones)", vbCritical, Err.Description
End Function

