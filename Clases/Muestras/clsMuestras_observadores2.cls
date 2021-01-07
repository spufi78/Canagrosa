VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestras_observadores2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMUESTRAS_OBSERVADORES
'***************************************************************
Private MUESTRA_ID As Long
Private DETERMINACION_ENSAYO_ID As Long
Private TIPO_DETERMINACION_ENSAYO_ID As Long
Private DOCUMENTO_ID As Long
Private Fecha As String
Private ES_CONTROL_EFICACIA As Integer

Private USUARIO_ID_REALIZACION As Long
Private USUARIO_ID_OBSERVADOR As Long
Private REALIZACION_ES_FORMADOR As Integer
Private OBSERVADOR_ES_FORMADOR As Integer

Private ES_DUPLICADO As Integer
Private ES_RECUALIFICACION As Integer
Private OBSERVACIONES As String


Private documento As String
Private USUARIO As String
Private USUARIO_OBSERVADOR As String
Private TIPO_OBSERVACION_ID As Long
Private TIPO_OBSERVACION As String

'Public Function Listado_Filtrado(Optional ByVal id_formador As Long = 0, _
'Optional ByVal id_usuario_en_formacion As Long = 0, _
'Optional ByVal id_pnt As Long = 0, _
'Optional ByVal id_tipo_muestra As Long = 0, _
'Optional ByVal FECHA_desde As String = "", _
'Optional ByVal FECHA_hasta As String = "", _
'Optional ByVal recual_todos As Boolean = True, _
'Optional ByVal ES_RECUALIFICACION As Boolean = True, _
'Optional ByVal formador_actor As Integer = 0, _
'Optional ByVal formador_observador As Integer = 0) As ADODB.RecordSet
'
'
'        Dim consulta As String, strCad As String
'
'        strCad = ""
'        If id_formador > 0 Then
'            strCad = strCad & " AND ((MUESTRAS_OBSERVADORES.USUARIO_ID_REALIZACION = " & CStr(id_formador) & " AND MUESTRAS_OBSERVADORES.REALIZACION_ES_FORMADOR = 1) OR (MUESTRAS_OBSERVADORES.USUARIO_ID_OBSERVADOR = " & CStr(id_formador) & " AND MUESTRAS_OBSERVADORES.REALIZACION_ES_FORMADOR = 0))"
'        End If
'
'        If id_usuario_en_formacion > 0 Then
'            strCad = strCad & " AND ((MUESTRAS_OBSERVADORES.USUARIO_ID_REALIZACION = " & CStr(id_usuario_en_formacion) & " AND MUESTRAS_OBSERVADORES.REALIZACION_ES_FORMADOR = 0) OR (MUESTRAS_OBSERVADORES.USUARIO_ID_OBSERVADOR = " & CStr(id_usuario_en_formacion) & " AND MUESTRAS_OBSERVADORES.REALIZACION_ES_FORMADOR = 1))"
'        End If
'
'        If id_pnt > 0 Then strCad = strCad & " AND MUESTRAS_OBSERVADORES.DOCUMENTO_ID = " & CStr(id_pnt)
'
'        If id_tipo_muestra > 0 Then strCad = strCad & " AND MUESTRAS.TIPO_MUESTRA_ID = " & CStr(id_tipo_muestra)
'
'        If Trim(FECHA_desde) <> "" Then strCad = strCad & " AND MUESTRAS_OBSERVADORES.FECHA >= '" & Format(CDate(FECHA_desde), "yyyy-mm-dd 00:00:00") & "'"
'
'        If Trim(FECHA_hasta) <> "" Then strCad = strCad & " AND MUESTRAS_OBSERVADORES.FECHA <= '" & Format(CDate(FECHA_hasta), "yyyy-mm-dd 23:59:59") & "'"
'
'        'If formador_actor = vbChecked Then strCad = strCad & " AND MUESTRAS_OBSERVADORES.REALIZACION_ES_FORMADOR = 1"
'
'        'If formador_observador = vbChecked Then strCad = strCad & " AND MUESTRAS_OBSERVADORES.REALIZACION_ES_FORMADOR = 0"
'
'        If Not recual_todos Then
'            strCad = strCad & " AND MUESTRAS_OBSERVADORES.ES_RECUALIFICACION = " & IIf(ES_RECUALIFICACION, "1", "0")
'        End If
'
'        consulta = "SELECT MUESTRAS_OBSERVADORES.*, CONCAT('[', COALESCE(CA_DOCUMENTOS.CODIGO, ''), '] ', COALESCE(CA_DOCUMENTOS.NOMBRE, 'N/A')) AS DOCUMENTO "
'        consulta = consulta & ", DECODIFICADORA.DESCRIPCION AS TIPO_OBSERVACION "
'        consulta = consulta & ", CONCAT(COALESCE(USUARIOS.NOMBRE, ''), ' ' , COALESCE(USUARIOS.APELLIDOS, '')) AS USUARIO"
'        consulta = consulta & ", CONCAT(COALESCE(USUARIOS_OB.NOMBRE, ''), ' ' , COALESCE(USUARIOS_OB.APELLIDOS, '')) AS USUARIO_OBSERVADOR"
'        consulta = consulta & ", CONCAT(tipos_muestra.CODIGO, '-', cast(muestras.ID_PARTICULAR as char(12))) as MUESTRA"
'        consulta = consulta & " FROM MUESTRAS_OBSERVADORES"
'        consulta = consulta & " LEFT OUTER JOIN CA_DOCUMENTOS ON MUESTRAS_OBSERVADORES.DOCUMENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
'        consulta = consulta & " LEFT OUTER JOIN USUARIOS ON MUESTRAS_OBSERVADORES.USUARIO_ID_REALIZACION = USUARIOS.ID_EMPLEADO"
'        consulta = consulta & " LEFT OUTER JOIN USUARIOS AS USUARIOS_OB ON MUESTRAS_OBSERVADORES.USUARIO_ID_OBSERVADOR = USUARIOS_OB.ID_EMPLEADO"
'        consulta = consulta & " INNER JOIN DECODIFICADORA ON MUESTRAS_OBSERVADORES.TIPO_OBSERVACION_ID = DECODIFICADORA.VALOR AND DECODIFICADORA.CODIGO = " & CStr(decodificadora.MC_TIPOS_OBSERVACION_ENSAYO)
'        consulta = consulta & " inner join muestras on muestras_observadores.MUESTRA_ID = muestras.ID_MUESTRA"
'        consulta = consulta & " inner join tipos_muestra on muestras.TIPO_MUESTRA_ID = tipos_muestra.ID_TIPO_MUESTRA"
'        consulta = consulta & " WHERE 1=1 " & strCad
'        consulta = consulta & " ORDER BY MUESTRA_ID, TIPO_DETERMINACION_ENSAYO_ID, DETERMINACION_ENSAYO_ID"
'
'        Set Listado_Filtrado = datos_bd(consulta)
'
'End Function

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMUESTRAS_OBSERVADORES
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setDETERMINACION_ENSAYO_ID(ByVal dato As Long)
        DETERMINACION_ENSAYO_ID = dato
End Property
Public Property Let setTIPO_DETERMINACION_ENSAYO_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ENSAYO_ID = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setDOCUMENTO(ByVal dato As String)
        documento = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        Fecha = dato
End Property
Public Property Let setES_CONTROL_EFICACIA(ByVal dato As Integer)
        ES_CONTROL_EFICACIA = dato
End Property
Public Property Let setTIPO_OBSERVACION_ID(ByVal dato As Long)
        TIPO_OBSERVACION_ID = dato
End Property
Public Property Let setTIPO_OBSERVACION(ByVal dato As String)
        TIPO_OBSERVACION = dato
End Property
Public Property Let setUSUARIO(ByVal dato As String)
        USUARIO = dato
End Property
Public Property Let setUSUARIO_ID_REALIZACION(ByVal dato As Long)
        USUARIO_ID_REALIZACION = dato
End Property

Public Property Let setUSUARIO_OBSERVADOR(ByVal dato As String)
        USUARIO_OBSERVADOR = dato
End Property

Public Property Let setUSUARIO_ID_OBSERVADOR(ByVal dato As Long)
        USUARIO_ID_OBSERVADOR = dato
End Property
Public Property Let setES_RECUALIFICACION(ByVal dato As Integer)
        ES_RECUALIFICACION = dato
End Property
Public Property Let setREALIZACION_ES_FORMADOR(ByVal dato As Integer)
        REALIZACION_ES_FORMADOR = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMUESTRAS_OBSERVADORES
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getDETERMINACION_ENSAYO_ID() As Long
        getDETERMINACION_ENSAYO_ID = DETERMINACION_ENSAYO_ID
End Property
Public Property Get getTIPO_DETERMINACION_ENSAYO_ID() As Long
        getTIPO_DETERMINACION_ENSAYO_ID = TIPO_DETERMINACION_ENSAYO_ID
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getDOCUMENTO() As String
        getDOCUMENTO = documento
End Property

Public Property Get getFECHA() As String
        getFECHA = Fecha
End Property
Public Property Get getES_CONTROL_EFICACIA() As Integer
        getES_CONTROL_EFICACIA = ES_CONTROL_EFICACIA
End Property
Public Property Get getUSUARIO_ID_REALIZACION() As Long
        getUSUARIO_ID_REALIZACION = USUARIO_ID_REALIZACION
End Property
Public Property Get getTIPO_OBSERVACION_ID() As Long
        getTIPO_OBSERVACION_ID = TIPO_OBSERVACION_ID
End Property

Public Property Get getTIPO_OBSERVACION() As String
        getTIPO_OBSERVACION = TIPO_OBSERVACION
End Property

Public Property Get getUSUARIO() As String
        getUSUARIO = USUARIO
End Property
Public Property Get getUSUARIO_OBSERVADOR() As String
        getUSUARIO_OBSERVADOR = USUARIO_OBSERVADOR
End Property

Public Property Get getUSUARIO_ID_OBSERVADOR() As Long
        getUSUARIO_ID_OBSERVADOR = USUARIO_ID_OBSERVADOR
End Property
Public Property Get getES_RECUALIFICACION() As Integer
        getES_RECUALIFICACION = ES_RECUALIFICACION
End Property
Public Property Get getREALIZACION_ES_FORMADOR() As Integer
        getREALIZACION_ES_FORMADOR = REALIZACION_ES_FORMADOR
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMUESTRAS_OBSERVADORES
'***************************************************************
Public Function Carga(ID_MUESTRA As Long, ID_DETERMINACION As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MUESTRAS_OBSERVADORES.*, CONCAT('[', COALESCE(CA_DOCUMENTOS.CODIGO, ''), '] ', COALESCE(CA_DOCUMENTOS.NOMBRE, 'N/A')) AS DOCUMENTO "
        consulta = consulta & ", DECODIFICADORA.DESCRIPCION AS TIPO_OBSERVACION "
        consulta = consulta & ", CONCAT(COALESCE(USUARIOS.NOMBRE, ''), ' ' , COALESCE(USUARIOS.APELLIDOS, '')) AS USUARIO"
        consulta = consulta & ", CONCAT(COALESCE(USUARIOS_OB.NOMBRE, ''), ' ' , COALESCE(USUARIOS_OB.APELLIDOS, '')) AS USUARIO_OBSERVADOR"
        consulta = consulta & " FROM MUESTRAS_OBSERVADORES"
        consulta = consulta & " LEFT OUTER JOIN CA_DOCUMENTOS ON MUESTRAS_OBSERVADORES.DOCUMENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " LEFT OUTER JOIN USUARIOS ON MUESTRAS_OBSERVADORES.USUARIO_ID_REALIZACION = USUARIOS.ID_EMPLEADO"
        consulta = consulta & " LEFT OUTER JOIN USUARIOS AS USUARIOS_OB ON MUESTRAS_OBSERVADORES.USUARIO_ID_OBSERVADOR = USUARIOS_OB.ID_EMPLEADO"
        consulta = consulta & " INNER JOIN DECODIFICADORA ON MUESTRAS_OBSERVADORES.TIPO_OBSERVACION_ID = DECODIFICADORA.VALOR AND DECODIFICADORA.CODIGO = " & CStr(decodificadora.MC_TIPOS_OBSERVACION_ENSAYO)
        consulta = consulta & " WHERE MUESTRAS_OBSERVADORES.MUESTRA_ID = " & ID_MUESTRA & " and MUESTRAS_OBSERVADORES.DETERMINACION_ENSAYO_ID = " & ID_DETERMINACION
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                DETERMINACION_ENSAYO_ID = rs("DETERMINACION_ENSAYO_ID")
                TIPO_DETERMINACION_ENSAYO_ID = rs("TIPO_DETERMINACION_ENSAYO_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                documento = rs("DOCUMENTO")
                Fecha = rs("FECHA")
                ES_CONTROL_EFICACIA = CInt(rs("ES_CONTROL_EFICACIA"))
                USUARIO_ID_REALIZACION = rs("USUARIO_ID_REALIZACION")
                USUARIO_ID_OBSERVADOR = rs("USUARIO_ID_OBSERVADOR")
                USUARIO = rs("USUARIO")
                USUARIO_OBSERVADOR = rs("USUARIO_OBSERVADOR")
                TIPO_OBSERVACION = rs("TIPO_OBSERVACION")
                TIPO_OBSERVACION_ID = rs("TIPO_OBSERVACION_ID")
                ES_RECUALIFICACION = CInt(rs("ES_RECUALIFICACION"))
                REALIZACION_ES_FORMADOR = CInt(rs("REALIZACION_ES_FORMADOR"))
                
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMuestras_observadores)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO MUESTRAS_OBSERVADORES (MUESTRA_ID, DETERMINACION_ENSAYO_ID, FECHA)" & _
                   "  VALUES (" & MUESTRA_ID & "," & DETERMINACION_ENSAYO_ID & ",'" & Fecha & "')"
                        
        execute_bd consulta
        Insertar = 1
        
        Modificar
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMuestras_observadores)", vbCritical, Err.Description
End Function
Public Function Modificar() As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE MUESTRAS_OBSERVADORES SET " & _
                        " DETERMINACION_ENSAYO_ID = " & DETERMINACION_ENSAYO_ID & "," & _
                        " TIPO_DETERMINACION_ENSAYO_ID = " & TIPO_DETERMINACION_ENSAYO_ID & "," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " FECHA = '" & Fecha & "'," & _
                        " ES_CONTROL_EFICACIA = " & ES_CONTROL_EFICACIA & "," & _
                        " USUARIO_ID_REALIZACION = " & USUARIO_ID_REALIZACION & "," & _
                        " USUARIO_ID_OBSERVADOR = " & USUARIO_ID_OBSERVADOR & "," & _
                        " TIPO_OBSERVACION_ID = " & TIPO_OBSERVACION_ID & "," & _
                        " ES_RECUALIFICACION = " & ES_RECUALIFICACION & "," & _
                        " REALIZACION_ES_FORMADOR = " & REALIZACION_ES_FORMADOR & "" & _
                " WHERE MUESTRA_ID = " & MUESTRA_ID & _
                " AND DETERMINACION_ENSAYO_ID = " & CStr(DETERMINACION_ENSAYO_ID)
                
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMuestras_observadores)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID_MUESTRA As Long, ID_DETERMINACION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MUESTRAS_OBSERVADORES " & _
                "    WHERE MUESTRA_ID = " & ID_MUESTRA & _
                "    AND DETERMINACION_ENSAYO_ID = " & ID_DETERMINACION
                
        execute_bd consulta
        Eliminar = True
        Exit Function
        
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMuestras_observadores)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Set Listado = Listado_Filtrado()
End Function

