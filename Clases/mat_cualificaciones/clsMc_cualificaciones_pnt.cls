VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMc_cualificaciones_pnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMC_CUALIFICACIONES_PNT
'***************************************************************
Private USUARIO_ID As Long
Private DOCUMENTO_ID As Long
Private ES_FORMADOR As Integer
Private DOCUMENTO As String
Private USUARIO As String

Public Function comprobar_usuario_descualificable(ID_USUARIO) As Boolean

Dim consulta As String
Dim rs As ADODB.RecordSet
On Error GoTo comprobar_usuario_descualificable_Error
    
        consulta = "SELECT * FROM MUESTRAS_OBSERVADORES "
        consulta = consulta & " WHERE USUARIO_ID = " & CStr(ID_USUARIO)
        consulta = consulta & " AND ACTOR_ES_FORMADOR = 1"
        consulta = consulta & " UNION "
        consulta = consulta & " SELECT * FROM MUESTRAS_OBSERVADORES "
        consulta = consulta & " WHERE USUARIO_OBSERVADOR_ID = " & CStr(ID_USUARIO)
        consulta = consulta & " AND ACTOR_ES_FORMADOR = 0"
                
        Set rs = datos_bd(consulta)
        
        comprobar_usuario_descualificable = (rs.RecordCount = 0)
            
        Set rs = Nothing
        
        
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsMc_cualificaciones_pnt.comprobar_usuario_descualificable"
    Exit Function
comprobar_usuario_descualificable_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsMc_cualificaciones_pnt.comprobar_usuario_descualificable"
    error_grave Err.Number & " (" & Err.Description & ") in procedure comprobar_usuario_descualificable of Módulo de clase clsMc_cualificaciones_pnt" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function

Public Function comprobar_usuario_es_cualificado(ID_USUARIO, id_pnt) As Boolean

Dim consulta As String
Dim rs As ADODB.RecordSet
On Error GoTo comprobar_usuario_es_cualificado_Error
    
        consulta = " SELECT * FROM MC_CUALIFICACIONES_PNT "
        consulta = consulta & " WHERE USUARIO_ID = " & CStr(ID_USUARIO)
        consulta = consulta & " AND DOCUMENTO_ID = " & CStr(id_pnt)
        consulta = consulta & " AND ES_FORMADOR = 1"
        
                
        Set rs = datos_bd(consulta)
        
        comprobar_usuario_es_cualificado = (rs.RecordCount <> 0)
            
        Set rs = Nothing
        
        
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsMc_cualificaciones_pnt.comprobar_usuario_descualificable"
    Exit Function
comprobar_usuario_es_cualificado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsMc_cualificaciones_pnt.comprobar_usuario_descualificable"
    error_grave Err.Number & " (" & Err.Description & ") in procedure comprobar_usuario_descualificable of Módulo de clase clsMc_cualificaciones_pnt" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function


Public Function comprobar_usuario_descualificable_en_pnt(ID_USUARIO, ID_DOC) As Boolean

Dim consulta As String
Dim rs As ADODB.RecordSet
On Error GoTo comprobar_usuario_descualificable_Error
    
        consulta = "SELECT * FROM MUESTRAS_OBSERVADORES "
        consulta = consulta & " WHERE USUARIO_ID = " & CStr(ID_USUARIO)
        consulta = consulta & " AND ACTOR_ES_FORMADOR = 1"
        consulta = consulta & " AND DOCUMENTO_ID = " & CStr(ID_DOC)
        consulta = consulta & " UNION "
        consulta = consulta & " SELECT * FROM MUESTRAS_OBSERVADORES "
        consulta = consulta & " WHERE USUARIO_OBSERVADOR_ID = " & CStr(ID_USUARIO)
        consulta = consulta & " AND ACTOR_ES_FORMADOR = 0"
        consulta = consulta & " AND DOCUMENTO_ID = " & CStr(ID_DOC)
                
        Set rs = datos_bd(consulta)
        
        comprobar_usuario_descualificable_en_pnt = (rs.RecordCount = 0)
            
        Set rs = Nothing
        
        
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsMc_cualificaciones_pnt.comprobar_usuario_descualificable"
    Exit Function
comprobar_usuario_descualificable_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsMc_cualificaciones_pnt.comprobar_usuario_descualificable"
    error_grave Err.Number & " (" & Err.Description & ") in procedure comprobar_usuario_descualificable of Módulo de clase clsMc_cualificaciones_pnt" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function

Public Function desmarcar_usuario(ByVal ID_USUARIO As Long, ByVal ID_DOCUMENTO As Long, Optional ByVal ID_USUARIO_CALIDAD As Long = 0) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MC_CUALIFICACIONES_PNT SET " & _
                        " ES_FORMADOR = 0" & _
                " WHERE USUARIO_ID = " & ID_USUARIO & _
                " AND DOCUMENTO_ID = " & ID_DOCUMENTO
                
        execute_bd consulta
        
        consulta = "UPDATE MC_CUALIFICACIONES_PNT_REGISTRO " & _
        " SET FECHA_BAJA = '" & Format(Now, "YYYY-MM-DD") & "'" & _
        " , USUARIO_CALIDAD_BAJA_ID = " & ID_USUARIO_CALIDAD & _
        " WHERE USUARIO_ID = " & ID_USUARIO & _
        " AND DOCUMENTO_ID = " & ID_DOCUMENTO & _
        " AND FECHA_ALTA <> '0000-00-00'" & _
        " AND FECHA_BAJA = '0000-00-00'"
        
        execute_bd consulta
        
        
        ' AHORA COMPRUEBA QUE NO ESTÉ EL MISMO DÍA LA ALTA Y LA BAJA
        consulta = "DELETE FROM MC_CUALIFICACIONES_PNT_REGISTRO "
        consulta = consulta & " WHERE USUARIO_ID = " & USUARIO_ID
        consulta = consulta & " AND DOCUMENTO_ID = " & DOCUMENTO_ID
        consulta = consulta & " AND FECHA_ALTA = FECHA_ALTA"
        execute_bd consulta
        
        
        
        desmarcar_usuario = True
        Exit Function
fallo:
        desmarcar_usuario = False
        MsgBox "Error al Desmarcar_Usuario (clsMc_cualificaciones_pnt)", vbCritical, Err.Description
End Function

Public Function Listado_por_Usuario(ByVal ID_USUARIO As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MC_CUALIFICACIONES_PNT.*  "
        consulta = consulta & ", CONCAT(COALESCE(USUARIOS.NOMBRE), ' ', COALESCE(USUARIOS.APELLIDOS)) AS USUARIO "
        consulta = consulta & ", CONCAT('[', COALESCE(CA_DOCUMENTOS.CODIGO), '] ', COALESCE(CA_DOCUMENTOS.NOMBRE)) AS DOCUMENTO "
        consulta = consulta & " FROM MC_CUALIFICACIONES_PNT"
        consulta = consulta & " LEFT OUTER JOIN USUARIOS ON MC_CUALIFICACIONES_PNT.USUARIO_ID = USUARIOS.ID_EMPLEADO"
        consulta = consulta & " LEFT OUTER JOIN CA_DOCUMENTOS ON MC_CUALIFICACIONES_PNT.DOCUMENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " WHERE MC_CUALIFICACIONES_PNT.USUARIO_ID = " & CStr(ID_USUARIO)
        consulta = consulta & " ORDER BY DOCUMENTO_ID"
        
        Set Listado_por_Usuario = datos_bd(consulta)

End Function

Public Function Listado_por_pnt(ByVal id_pnt As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MC_CUALIFICACIONES_PNT.*  "
        consulta = consulta & ", CONCAT(COALESCE(USUARIOS.NOMBRE), ' ', COALESCE(USUARIOS.APELLIDOS)) AS USUARIO "
        consulta = consulta & ", CONCAT('[', COALESCE(CA_DOCUMENTOS.CODIGO), '] ', COALESCE(CA_DOCUMENTOS.NOMBRE)) AS DOCUMENTO "
        consulta = consulta & " FROM MC_CUALIFICACIONES_PNT"
        consulta = consulta & " LEFT OUTER JOIN USUARIOS ON MC_CUALIFICACIONES_PNT.USUARIO_ID = USUARIOS.ID_EMPLEADO"
        consulta = consulta & " LEFT OUTER JOIN CA_DOCUMENTOS ON MC_CUALIFICACIONES_PNT.DOCUMENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " WHERE MC_CUALIFICACIONES_PNT.DOCUMENTO_ID = " & CStr(id_pnt)
        consulta = consulta & " ORDER BY USUARIO_ID"
        
        Set Listado_por_pnt = datos_bd(consulta)

End Function


Public Function Listado_usuarios_cualificados(Optional ByVal solo_cualificados_actualmente As Boolean = False) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT USUARIOS.* FROM USUARIOS"
        consulta = consulta & " WHERE USUARIOS.ID_EMPLEADO IN (SELECT MC_CUALIFICACIONES_PNT.USUARIO_ID FROM MC_CUALIFICACIONES_PNT"
        If solo_cualificados_actualmente Then
            consulta = consulta & " WHERE MC_CUALIFICACIONES_PNT.ES_FORMADOR = 1"
        End If
        consulta = consulta & ")"
        
        
        Set Listado_usuarios_cualificados = datos_bd(consulta)

End Function

Public Function llenar_combo_usuarios_cualificados(ByRef combo As miCombo, id_usuario_actual, Optional ByVal solo_cualificados_actualmente As Boolean = True) As Boolean

    Dim ousu As New clsUsuarios
    Dim strFiltro As String
    
    strFiltro = " AND id_empleado <> " & id_usuario_actual
    strFiltro = strFiltro & " AND ID_EMPLEADO IN (SELECT MC_CUALIFICACIONES_PNT.USUARIO_ID FROM MC_CUALIFICACIONES_PNT"
    If solo_cualificados_actualmente Then
        strFiltro = strFiltro & " WHERE MC_CUALIFICACIONES_PNT.ES_FORMADOR = 1 "
    End If
    strFiltro = strFiltro & ")"
    

    ousu.llenar_combo combo, 0, frmUsuarios, strFiltro
    
    Set ousu = Nothing

End Function


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMC_CUALIFICACIONES_PNT
'***************************************************************
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setES_FORMADOR(ByVal dato As Integer)
        ES_FORMADOR = dato
End Property
Public Property Let setUSUARIO(ByVal dato As String)
        USUARIO = dato
End Property
Public Property Let setDOCUMENTO(ByVal dato As String)
        DOCUMENTO = dato
End Property


'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMC_CUALIFICACIONES_PNT
'***************************************************************
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getES_FORMADOR() As Integer
        getES_FORMADOR = ES_FORMADOR
End Property
Public Property Get getDOCUMENTO() As Long
        getDOCUMENTO = DOCUMENTO
End Property
Public Property Get getUSUARIO() As Integer
        getUSUARIO = USUARIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMC_CUALIFICACIONES_PNT
'***************************************************************
Public Function Carga_PNT(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MC_CUALIFICACIONES_PNT.*  "
        consulta = consulta & ", CONCAT(COALESCE(USUARIOS.NOMBRE), ' ', COALESCE(USUARIOS.APELLIDOS)) AS USUARIO "
        consulta = consulta & ", CONCAT('[', COALESCE(CA_DOCUMENTOS.CODIGO), '] ', COALESCE(CA_DOCUMENTOS.NOMBRE)) AS DOCUMENTO "
        consulta = consulta & " FROM MC_CUALIFICACIONES_PNT"
        consulta = consulta & " LEFT OUTER JOIN USUARIOS ON MC_CUALIFICACIONES_PNT.USUARIO_ID = USUARIOS.ID_EMPLEADO"
        consulta = consulta & " LEFT OUTER JOIN CA_DOCUMENTOS ON MC_CUALIFICACIONES_PNT.DOCUMENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " WHERE MC_CUALIFICACIONES_PNT.DOCUMENTO_ID = " & CStr(ID)
        consulta = consulta & " ORDER BY USUARIO_ID"
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_PNT = False
        Else
                USUARIO_ID = rs("USUARIO_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                USUARIO = rs("USUARIO")
                DOCUMENTO = rs("DOCUMENTO")
                ES_FORMADOR = rs("ES_FORMADOR")
                Carga_PNT = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_PNT = False
        MsgBox "Error al cargar los datos por PNT(clsMc_cualificaciones_pnt)", vbCritical, Err.Description
End Function

Public Function Carga_USUARIO(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MC_CUALIFICACIONES_PNT.*  "
        consulta = consulta & ", CONCAT(COALESCE(USUARIOS.NOMBRE), ' ', COALESCE(USUARIOS.APELLIDOS)) AS USUARIO "
        consulta = consulta & ", CONCAT('[', COALESCE(CA_DOCUMENTOS.CODIGO), '] ', COALESCE(CA_DOCUMENTOS.NOMBRE)) AS DOCUMENTO "
        consulta = consulta & " FROM MC_CUALIFICACIONES_PNT"
        consulta = consulta & " LEFT OUTER JOIN USUARIOS ON MC_CUALIFICACIONES_PNT.USUARIO_ID = USUARIOS.ID_EMPLEADO"
        consulta = consulta & " LEFT OUTER JOIN CA_DOCUMENTOS ON MC_CUALIFICACIONES_PNT.DOCUMENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " WHERE MC_CUALIFICACIONES_PNT.USUARIO_ID = " & CStr(ID)
        consulta = consulta & " ORDER BY DOCUMENTO_ID"

        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_USUARIO = False
        Else
                USUARIO_ID = rs("USUARIO_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                USUARIO = rs("USUARIO")
                DOCUMENTO = rs("DOCUMENTO")
                ES_FORMADOR = rs("ES_FORMADOR")
                Carga_USUARIO = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_USUARIO = False
        MsgBox "Error al cargar los datos por USUARIO(clsMc_cualificaciones_pnt)", vbCritical, Err.Description
End Function

Public Function Carga(ID_USUARIO As Long, ID_DOCUMENTO) As Boolean
    On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        
        Dim consulta As String
        
        consulta = "SELECT MC_CUALIFICACIONES_PNT.*"
        consulta = consulta & ", CONCAT(COALESCE(USUARIOS.NOMBRE), ' ', COALESCE(USUARIOS.APELLIDOS)) AS USUARIO "
        consulta = consulta & ", CONCAT('[', COALESCE(CA_DOCUMENTOS.CODIGO), '] ', COALESCE(CA_DOCUMENTOS.NOMBRE)) AS DOCUMENTO "
        consulta = consulta & " FROM MC_CUALIFICACIONES_PNT"
        consulta = consulta & " LEFT OUTER JOIN USUARIOS ON MC_CUALIFICACIONES_PNT.USUARIO_ID = USUARIOS.ID_EMPLEADO"
        consulta = consulta & " LEFT OUTER JOIN CA_DOCUMENTOS ON MC_CUALIFICACIONES_PNT.DOCUMENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " WHERE MC_CUALIFICACIONES_PNT.USUARIO_ID = " & CStr(ID_USUARIO)
        consulta = consulta & " AND MC_CUALIFICACIONES_PNT.DOCUMENTO_ID = " & CStr(ID_DOCUMENTO)

        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                USUARIO_ID = rs("USUARIO_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                USUARIO = rs("USUARIO")
                DOCUMENTO = rs("DOCUMENTO")
                ES_FORMADOR = rs("ES_FORMADOR")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        Set rs = Nothing
        MsgBox "Error al cargar los datos por USUARIO(clsMc_cualificaciones_pnt)", vbCritical, Err.Description
End Function

Public Function Insertar(Optional ByVal ID_USUARIO_CALIDAD As Long = 0) As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO MC_CUALIFICACIONES_PNT " & _
                   "  VALUES (" & _
                        USUARIO_ID & "," & DOCUMENTO_ID & "," & ES_FORMADOR & _
                ")"
        execute_bd consulta
        Insertar = USUARIO_ID
        
        ' INSERTA REGISTRO DE ALTA
        consulta = "insert into mc_cualificaciones_pnt_registro (usuario_id, documento_id, fecha_alta, fecha_baja, usuario_calidad_alta_id, usuario_calidad_baja_id) "
        consulta = consulta & " values (" & USUARIO_ID
        consulta = consulta & ", " & DOCUMENTO_ID
        consulta = consulta & ", '" & Format(Now, "YYYY-MM-DD") & "'"
        consulta = consulta & ", '0000-00-00'"
        consulta = consulta & ", " & ID_USUARIO_CALIDAD
        consulta = consulta & ", 0"
        consulta = consulta & ")"
        
        execute_bd consulta
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMc_cualificaciones_pnt)", vbCritical, Err.Description
End Function

Public Function Modificar(Optional ByVal ID_USUARIO_CALIDAD As Long = 0) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MC_CUALIFICACIONES_PNT SET " & _
                        " ES_FORMADOR = " & ES_FORMADOR & "" & _
                " WHERE USUARIO_ID = " & USUARIO_ID & _
                " AND DOCUMENTO_ID = " & DOCUMENTO_ID
                
        execute_bd consulta
        
        ' Ahora comprueba si es alta o baja.
        If ES_FORMADOR = 1 Then
                ' INSERTA REGISTRO DE ALTA
            consulta = "insert into mc_cualificaciones_pnt_registro (usuario_id, documento_id, fecha_alta, fecha_baja, usuario_calidad_alta_id, usuario_calidad_baja_id) "
            consulta = consulta & " values (" & USUARIO_ID
            consulta = consulta & ", " & DOCUMENTO_ID
            consulta = consulta & ", '" & Format(Now, "YYYY-MM-DD") & "'"
            consulta = consulta & ", '0000-00-00'"
            consulta = consulta & ", " & ID_USUARIO_CALIDAD
            consulta = consulta & ", 0"
            consulta = consulta & ")"
        Else
            consulta = "UPDATE mc_cualificaciones_pnt_registro SET"
            consulta = consulta & ", FECHA_BAJA = '" & Format(Now, "YYYY-MM-DD") & "'"
            consulta = consulta & ", USUARIO_CALIDAD_BAJA_ID = " & ID_USUARIO_CALIDAD
            consulta = consulta & " WHERE USUARIO_ID = " & USUARIO_ID
            consulta = consulta & " AND DOCUMENTO_ID = " & USUARIO_ID
            consulta = consulta & " AND FECHA_ALTA <> '0000-00-00'"
            consulta = consulta & " AND FECHA_BAJA = '0000-00-00'"
            
            execute_bd consulta
            
            ' AHORA COMPRUEBA QUE NO ESTÉ EL MISMO DÍA LA ALTA Y LA BAJA
            consulta = "DELETE FROM MC_CUALIFICACIONES_PNT_REGISTRO "
            consulta = consulta & " WHERE USUARIO_ID = " & USUARIO_ID
            consulta = consulta & " AND DOCUMENTO_ID = " & DOCUMENTO_ID
            consulta = consulta & " AND FECHA_ALTA = FECHA_ALTA"
            
        End If
        
        execute_bd consulta
        
        
        
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMc_cualificaciones_pnt)", vbCritical, Err.Description
End Function
Public Function Eliminar(id_usu As Long, ID_DOC As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        ' PRIMERO ELIMINA EL REGISTRO MAS RECIENTE DEL REGISTRO
        
        consulta = "DELETE FROM MC_CUALIFICACIONES_PNT_REGISTRO " & _
        " WHERE USUARIO_ID = " & id_usu & _
        " AND DOCUMENTO_ID = " & ID_DOC & _
        " AND FECHA_ALTA <> '0000-00-00'" & _
        " AND FECHA_BAJA = '0000-00-00'"
        
        execute_bd consulta
        
        
        consulta = "DELETE FROM MC_CUALIFICACIONES_PNT " & _
                "    WHERE USUARIO_ID = " & id_usu & _
                "    AND DOCUMENTO_ID = " & ID_DOC
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMc_cualificaciones_pnt)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MC_CUALIFICACIONES_PNT.*  "
        consulta = consulta & ", CONCAT(COALESCE(USUARIOS.NOMBRE), ' ', COALESCE(USUARIOS.APELLIDOS)) AS USUARIO "
        consulta = consulta & ", CONCAT('[', COALESCE(CA_DOCUMENTOS.CODIGO), '] ', COALESCE(CA_DOCUMENTOS.NOMBRE)) AS DOCUMENTO "
        consulta = consulta & " FROM MC_CUALIFICACIONES_PNT"
        consulta = consulta & " LEFT OUTER JOIN USUARIOS ON MC_CUALIFICACIONES_PNT.USUARIO_ID = USUARIOS.ID_EMPLEADO"
        consulta = consulta & " LEFT OUTER JOIN CA_DOCUMENTOS ON MC_CUALIFICACIONES_PNT.DOCUMENTO_ID = CA_DOCUMENTOS.ID_DOCUMENTO"
        consulta = consulta & " ORDER BY MC_CUALIFICACIONES_PNT.DOCUMENTO_ID, MC_CUALIFICACIONES_PNT.USUARIO_ID"
        
        Set Listado = datos_bd(consulta)
End Function

