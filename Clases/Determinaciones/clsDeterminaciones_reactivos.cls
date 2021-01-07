VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeterminaciones_reactivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDETERMINACIONES_REACTIVOS
'***************************************************************
Private DETERMINACION_ID As Long
Private BOTE_EX_ID As Long
Private tipo As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDETERMINACIONES_REACTIVOS
'***************************************************************
Public Property Let setDETERMINACION_ID(ByVal dato As Long)
        DETERMINACION_ID = dato
End Property
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As String)
        tipo = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDETERMINACIONES_REACTIVOS
'***************************************************************
Public Property Get getDETERMINACION_ID() As Long
        getDETERMINACION_ID = DETERMINACION_ID
End Property
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getTIPO() As String
        getTIPO = tipo
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDETERMINACIONES_REACTIVOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DETERMINACIONES_REACTIVOS WHERE DETERMINACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DETERMINACION_ID = rs("DETERMINACION_ID")
                BOTE_EX_ID = rs("BOTE_EX_ID")
                tipo = rs("TIPO")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDeterminaciones_reactivos)", vbCritical, Err.Description
End Function
Public Function Insertar(lDETERMINACION As Long, lTIPO_DETERMINACION As Long) As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = " INSERT INTO DETERMINACIONES_REACTIVOS " & _
                   " SELECT " & lDETERMINACION & ",BOTE_EX_ID,TIPO,ORDEN " & _
                   "   FROM TIPOS_DETERMINACION_BOTES_EX " & _
                   "  WHERE TIPO_DETERMINACION_ID = " & lTIPO_DETERMINACION
        execute_bd consulta
        Insertar = lDETERMINACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDeterminaciones_equipos)", vbCritical, Err.Description
End Function
Public Function Insertar_Determinacion() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = " INSERT INTO DETERMINACIONES_REACTIVOS VALUES (" & _
                     DETERMINACION_ID & "," & BOTE_EX_ID & ",'" & tipo & "'," & ORDEN & ")"
        execute_bd consulta
        Insertar_Determinacion = DETERMINACION_ID
        Exit Function
fallo:
        Insertar_Determinacion = 0
        MsgBox "Error al Insertar_Determinacion (clsDeterminaciones_equipos)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DETERMINACIONES_REACTIVOS SET " & _
                        " BOTE_EX_ID = " & BOTE_EX_ID & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE DETERMINACION_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDeterminaciones_reactivos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DETERMINACIONES_REACTIVOS " & _
                "    WHERE DETERMINACION_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDeterminaciones_reactivos)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.BOTE_EX_ID,A.TIPO " & _
                   " FROM DETERMINACIONES_REACTIVOS A " & _
                   " WHERE A.DETERMINACION_ID = " & ID & _
                   " ORDER BY A.ORDEN"
        
'        consulta = "SELECT B.ID_BOTE_EX,D.NOMBRE,B.FECHA_CADUCIDAD " & _
'                   " FROM DETERMINACIONES_REACTIVOS A, BOTES_EX B, TIPOS_BOTE_EX C, TIPOS_REACTIVO_EX D " & _
'                   " WHERE A.DETERMINACION_ID = " & ID & _
'                   "   AND A.BOTE_EX_ID = B.ID_BOTE_EX " & _
'                   "   AND B.TIPO_BOTE_EX_ID = C.ID_TIPO_BOTE_EX " & _
'                   "   AND C.TIPO_REACTIVO_EX_ID = D.ID_TIPO_REACTIVO_EX " & _
'                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
Public Function duplicar(ID_DETERMINACION_ORIGEN As Long, ID_DETERMINACION_DESTINO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE INTO DETERMINACIONES_REACTIVOS " & _
                   "SELECT " & ID_DETERMINACION_DESTINO & ",BOTE_EX_ID,TIPO,ORDEN FROM DETERMINACIONES_REACTIVOS WHERE DETERMINACION_ID = " & ID_DETERMINACION_ORIGEN
        execute_bd consulta
        duplicar = True
        Exit Function
fallo:
        duplicar = False
        MsgBox "Error al duplicar (clsDeterminaciones_reactivos)", vbCritical, Err.Description
End Function


