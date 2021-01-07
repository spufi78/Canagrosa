VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVideos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSVIDEOS
'***************************************************************
Private ID_VIDEO As Long
Private DESCRIPCION As String
Private fecha As Date
Private OBSERVACIONES As String
Private TIPO_VIDEO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSVIDEOS
'***************************************************************
Public Property Let setID_VIDEO(ByVal dato As Long)
        ID_VIDEO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFECHA(ByVal dato As Date)
        fecha = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setTIPO_VIDEO_ID(ByVal dato As Long)
        TIPO_VIDEO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSVIDEOS
'***************************************************************
Public Property Get getID_VIDEO() As Long
        getID_VIDEO = ID_VIDEO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA() As Date
        getFECHA = fecha
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getTIPO_VIDEO_ID() As Long
        getTIPO_VIDEO_ID = TIPO_VIDEO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSVIDEOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM VIDEOS WHERE ID_VIDEO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_VIDEO = rs("ID_VIDEO")
                DESCRIPCION = rs("DESCRIPCION")
                fecha = Format(rs("FECHA"), "yyyy-mm-dd")
                OBSERVACIONES = rs("OBSERVACIONES")
                TIPO_VIDEO_ID = rs("TIPO_VIDEO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsVideos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_VIDEO) FROM VIDEOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_VIDEO = 1
        Else
                ID_VIDEO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO VIDEOS (ID_VIDEO, DESCRIPCION, FECHA, OBSERVACIONES, TIPO_VIDEO_ID)" & _
                   "  VALUES (" & _
                        ID_VIDEO & "," & "'" & DESCRIPCION & "'" & "," & "'" & Format(fecha, "yyyy-mm-dd") & "'" & "," & _
                        "'" & OBSERVACIONES & "'" & "," & TIPO_VIDEO_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_VIDEO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsVideos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE VIDEOS SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FECHA = '" & Format(fecha, "yyyy-mm-dd") & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " TIPO_VIDEO_ID = " & TIPO_VIDEO_ID & _
                " WHERE ID_VIDEO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsVideos)", vbCritical, Err.Description
End Function
Public Function Comprobar_es_vinculado(ID As Long) As Boolean
    Comprobar_es_vinculado = False
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM VIDEOS_DETALLE " & _
                "    WHERE VIDEO_ID = " & ID
        execute_bd consulta
        
        consulta = "DELETE FROM VIDEOS " & _
                "    WHERE ID_VIDEO = " & ID
        execute_bd consulta
        
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsVideos)", vbCritical, Err.Description
End Function
Public Function Listado(Optional ByVal prm_id_tipo_video As Long = 0, Optional ByVal prm_descripcion As String = "", Optional ByVal prm_desde As Date = Empty, Optional ByVal prm_hasta As Date = Empty) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        
        filtro = ""
        If prm_id_tipo_video > 0 Then
            filtro = filtro & " AND VIDEOS.TIPO_VIDEO_ID = " & CStr(prm_id_tipo_video)
        End If
                
        If Trim(prm_descripcion) <> "" Then
            filtro = filtro & " AND lower(concat(VIDEOS.descripcion, VIDEOS.observaciones)) like '%" & Trim(prm_descripcion) & "%'"
        End If
        
        If Format(prm_desde, "yyyy-mm-dd") <> Format(Now, "yyyy-mm-dd") Then
            filtro = filtro & " AND VIDEOS.fecha >= '" & Format(prm_desde, "yyyy-mm-dd") & "'"
        End If
        
        If Format(prm_hasta, "yyyy-mm-dd") <> Format(Now, "yyyy-mm-dd") Then
            filtro = filtro & " AND VIDEOS.fecha <= '" & Format(prm_hasta, "yyyy-mm-dd") & "'"
        End If
        
        consulta = "SELECT VIDEOS.*, DECODIFICADORA.DESCRIPCION as TIPO_VIDEO, coalesce(total_videos.total,0) as total_capitulos"
        consulta = consulta & " FROM VIDEOS"
        consulta = consulta & " INNER JOIN decodificadora ON VIDEOS.TIPO_VIDEO_ID = DECODIFICADORA.VALOR AND DECODIFICADORA.CODIGO = " & DECODIFICADORA.VIDEOS_TIPO_VIDEO
        consulta = consulta & " left outer join (select coalesce(video_id,-1) as video_id, count(orden) as total from videos_detalle) as total_videos on videos.id_video = total_videos.video_id"
        If Trim(filtro) <> "" Then
            consulta = consulta & " WHERE 1=1 " & filtro
        End If
        consulta = consulta & " ORDER BY VIDEOS.ID_VIDEO"
        
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_VIDEO,DESCRIPCION FROM VIDEOS ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
