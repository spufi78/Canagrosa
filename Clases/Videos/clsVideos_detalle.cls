VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVideos_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSVIDEOS_DETALLE
'***************************************************************
Private VIDEO_ID As Long
Private ORDEN As Long
Private DESCRIPCION As String
Private RUTA As String
Private OBSERVACIONES As String
Public Function comprobar_ruta_existente(ID As Long, RUTA As String) As Boolean

    Dim consulta As String, rs As ADODB.RecordSet
    
    
    consulta = "select * from videos_detalle where video_id = " & CStr(ID) & " and lower(ruta) = '" & LCase(Replace(RUTA, "\", "/")) & "'"
    
    Set rs = datos_bd(consulta)
    
    comprobar_ruta_existente = (rs.RecordCount <> 0)

    Set rs = Nothing
    
End Function

Public Function modificar_orden(ByVal prm_direccion As Integer, ByVal prm_orden As Long, ID As Long) As Boolean

Dim nuevo_orden As Long, consulta As String
Dim maxmin As Long, rs As ADODB.RecordSet

    If prm_direccion < 0 Then
        nuevo_orden = prm_orden - 1
        If prm_orden = 1 Then
            ' está en el límite para bajar orden
            Exit Function
        End If
    Else
        Set rs = datos_bd("select max(orden) as max_orden from videos_detalle where video_id = " & ID)
    
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            If Not IsNull(rs("max_orden")) Then
                maxmin = rs("max_orden")
            Else
                maxmin = 0
            End If
        End If
        
        nuevo_orden = prm_orden + 1
        If prm_orden = maxmin Then
            ' está en el límite para subir orden
            Exit Function
        End If

    End If

    consulta = "update videos_detalle set orden = " & CStr((-1) * nuevo_orden) & " where video_id = " & ID & " and orden = " & CStr(nuevo_orden)
    execute_bd consulta
    consulta = "update videos_detalle set orden = " & CStr(nuevo_orden) & " where video_id = " & ID & " and orden = " & CStr(prm_orden)
    execute_bd consulta
    consulta = "update videos_detalle set orden = " & CStr(prm_orden) & " where video_id = " & ID & " and orden = " & CStr((-1) * nuevo_orden)
    execute_bd consulta
    
End Function

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSVIDEOS_DETALLE
'***************************************************************
Public Property Let setVIDEO_ID(ByVal dato As Long)
        VIDEO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        RUTA = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSVIDEOS_DETALLE
'***************************************************************
Public Property Get getVIDEO_ID() As Long
        getVIDEO_ID = VIDEO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getRUTA() As String
        getRUTA = RUTA
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSVIDEOS_DETALLE
'***************************************************************
Public Function Carga(ID As Long, N_ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM VIDEOS_DETALLE WHERE VIDEO_ID = " & ID & " AND ORDEN = " & N_ORDEN
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                VIDEO_ID = rs("VIDEO_ID")
                ORDEN = rs("ORDEN")
                DESCRIPCION = rs("DESCRIPCION")
                RUTA = rs("RUTA")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsVideos_detalle)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ORDEN) FROM VIDEOS_DETALLE WHERE VIDEO_ID = " & VIDEO_ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ORDEN = 1
        Else
                ORDEN = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO VIDEOS_DETALLE " & _
                   "  VALUES (" & _
                        VIDEO_ID & "," & ORDEN & "," & "'" & DESCRIPCION & "'" & "," & _
                        "'" & RUTA & "'" & "," & "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = ORDEN
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsVideos_detalle)", vbCritical, Err.Description
End Function

Public Function Modificar() As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE VIDEOS_DETALLE SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " RUTA = '" & RUTA & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE VIDEO_ID = " & VIDEO_ID & " AND ORDEN = " & ORDEN
        
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsVideos_detalle)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, N_ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String, rs As ADODB.RecordSet, strCad As String
        
        consulta = "DELETE FROM VIDEOS_DETALLE " & _
                "    WHERE VIDEO_ID = " & ID & _
                "    AND ORDEN = " & N_ORDEN
              
        execute_bd consulta
        
        ' Ahora Comprueba el orden de los posteriores
        strCad = "SELECT * FROM VIDEOS_DETALLE WHERE VIDEO_ID = " & ID & " AND ORDEN = "
        
        Set rs = datos_bd(strCad & CStr(N_ORDEN + 1))
        While rs.RecordCount <> 0
            consulta = "update videos_detalle set orden = orden-1 where video_id = " & ID & " and orden = " & N_ORDEN + 1
            execute_bd consulta
            N_ORDEN = N_ORDEN + 1
            Set rs = datos_bd(strCad & CStr(N_ORDEN + 1))
        Wend
        
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsVideos_detalle)", vbCritical, Err.Description
End Function

Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT VIDEOS_DETALLE.*, coalesce(VIDEOS.DESCRIPCION, '') AS VIDEO, coalesce(DECODIFICADORA.DESCRIPCION, '') AS TIPO_VIDEO "
        consulta = consulta & " FROM VIDEOS_DETALLE "
        consulta = consulta & " left outer join videos on videos_detalle.video_id = videos.id_video "
        consulta = consulta & " left outer join decodificadora on videos.tipo_video_id = decodificadora.valor and decodificadora.codigo = " & decodificadora.VIDEOS_TIPO_VIDEO
        consulta = consulta & " ORDER BY VIDEO_ID"
        
        Set Listado = datos_bd(consulta)
End Function

Public Function Listado_detalle(ByVal ID_VIDEO As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT VIDEOS_DETALLE.*, coalesce(VIDEOS.DESCRIPCION, '') AS VIDEO, coalesce(DECODIFICADORA.DESCRIPCION, '') AS TIPO_VIDEO "
        consulta = consulta & " FROM VIDEOS_DETALLE "
        consulta = consulta & " left outer join videos on videos_detalle.video_id = videos.id_video "
        consulta = consulta & " left outer join decodificadora on videos.tipo_video_id = decodificadora.valor and decodificadora.codigo = " & decodificadora.VIDEOS_TIPO_VIDEO
        consulta = consulta & " WHERE videos_detalle.video_id = " & ID_VIDEO
        consulta = consulta & " ORDER BY VIDEO_ID"
        
        Set Listado_detalle = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT VIDEO_ID,DESCRIPCION FROM VIDEOS_DETALLE ORDER BY ORDEN"
        Set Listado_Combo = datos_bd(consulta)
End Function
