VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGranulometrias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_GRANULOMETRIA As Long
Private MUESTRA_ID As Long
Private DETERMINACION_ID As Long
Private PESO_100 As String
Private PARTIDOS As String
Private MERMADOS As String
Private OTROS_CEREALES As String
Private TRIGO_BLANDO As String
Private ATACADOS As String
Private CON_COLORACION As String
Private MACULADOS As String
Private FUSARIADOS As String
Private GERMINADOS As String
Private SEM_NOCIVAS As String
Private SEM_NO_NOCIVAS As String
Private AVERIADOS As String
Private OTROS_DANADOS As String
Private PMD As String
Private INSECTOS As String
Private VITREOS As String
Private BERRENDOS As String
Private HARINOSOS As String
Private PESO_MUESTRA As String
Private OBSERVACIONES_DET As String
Private ES_DUPLICADO As Integer
Public Property Let setID_GRANULOMETRIA(ByVal dato As Long)
    ID_GRANULOMETRIA = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
    MUESTRA_ID = dato
End Property
Public Property Let setDETERMINACION_ID(ByVal dato As Long)
    DETERMINACION_ID = dato
End Property
Public Property Let setPESO_100(ByVal dato As String)
    PESO_100 = dato
End Property
Public Property Let setPARTIDOS(ByVal dato As String)
    PARTIDOS = dato
End Property
Public Property Let setMERMADOS(ByVal dato As String)
    MERMADOS = dato
End Property
Public Property Let setOTROS_CEREALES(ByVal dato As String)
    OTROS_CEREALES = dato
End Property
Public Property Let setTRIGO_BLANDO(ByVal dato As String)
    TRIGO_BLANDO = Trim(dato)
End Property
Public Property Let setATACADOS(ByVal dato As String)
    ATACADOS = Trim(dato)
End Property
Public Property Let setCON_COLORACION(ByVal dato As String)
    CON_COLORACION = Trim(dato)
End Property
Public Property Let setMACULADOS(ByVal dato As String)
    MACULADOS = Trim(dato)
End Property
Public Property Let setFUSARIADOS(ByVal dato As String)
    FUSARIADOS = dato
End Property
Public Property Let setGERMINADOS(ByVal dato As String)
    GERMINADOS = dato
End Property
Public Property Let setSEM_NOCIVAS(ByVal dato As String)
    SEM_NOCIVAS = dato
End Property
Public Property Let setSEM_NO_NOCIVAS(ByVal dato As String)
    SEM_NO_NOCIVAS = dato
End Property
Public Property Let setAVERIADOS(ByVal dato As String)
    AVERIADOS = dato
End Property
Public Property Let setOTROS_DANADOS(ByVal dato As String)
    OTROS_DANADOS = dato
End Property
Public Property Let setPMD(ByVal dato As String)
    PMD = dato
End Property
Public Property Let setINSECTOS(ByVal dato As String)
    INSECTOS = dato
End Property
Public Property Let setVITREOS(ByVal dato As String)
    VITREOS = dato
End Property
Public Property Let setBERRENDOS(ByVal dato As String)
    BERRENDOS = dato
End Property
Public Property Let setHARINOSOS(ByVal dato As String)
    HARINOSOS = dato
End Property
Public Property Let setPESO_MUESTRA(ByVal dato As String)
    PESO_MUESTRA = dato
End Property
Public Property Let setOBSERVACIONES_DET(ByVal dato As String)
    OBSERVACIONES_DET = dato
End Property
Public Property Let setES_DUPLICADO(ByVal dato As Integer)
    ES_DUPLICADO = dato
End Property
Public Property Get getID_GRANULOMETRIA() As Long
    getID_GRANULOMETRIA = ID_GRANULOMETRIA
End Property
Public Property Get getMUESTRA_ID() As Long
    getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getDETERMINACION_ID() As Long
    getDETERMINACION_ID = DETERMINACION_ID
End Property
Public Property Get getPESO_100() As String
    getPESO_100 = PESO_100
End Property
Public Property Get getPARTIDOS() As String
    getPARTIDOS = PARTIDOS
End Property
Public Property Get getMERMADOS() As String
    getMERMADOS = MERMADOS
End Property
Public Property Get getOTROS_CEREALES() As String
    getOTROS_CEREALES = OTROS_CEREALES
End Property
Public Property Get getTRIGO_BLANDO() As String
    getTRIGO_BLANDO = TRIGO_BLANDO
End Property
Public Property Get getATACADOS() As String
    getATACADOS = ATACADOS
End Property
Public Property Get getCON_COLORACION() As String
    getCON_COLORACION = CON_COLORACION
End Property
Public Property Get getMACULADOS() As String
    getMACULADOS = MACULADOS
End Property
Public Property Get getFUSARIADOS() As String
    getFUSARIADOS = FUSARIADOS
End Property
Public Property Get getGERMINADOS() As String
    getGERMINADOS = GERMINADOS
End Property
Public Property Get getSEM_NOCIVAS() As String
    getSEM_NOCIVAS = SEM_NOCIVAS
End Property

Public Property Get getSEM_NO_NOCIVAS() As String
    getSEM_NO_NOCIVAS = SEM_NO_NOCIVAS
End Property
Public Property Get getAVERIADOS() As String
    getAVERIADOS = AVERIADOS
End Property
Public Property Get getOTROS_DANADOS() As String
    getOTROS_DANADOS = OTROS_DANADOS
End Property
Public Property Get getPMD() As String
    getPMD = PMD
End Property
Public Property Get getINSECTOS() As String
    getINSECTOS = INSECTOS
End Property
Public Property Get getVITREOS() As String
    getVITREOS = VITREOS
End Property
Public Property Get getBERRENDOS() As String
    getBERRENDOS = BERRENDOS
End Property
Public Property Get getHARINOSOS() As String
    getHARINOSOS = HARINOSOS
End Property
Public Property Get getPESO_MUESTRA() As String
    getPESO_MUESTRA = PESO_MUESTRA
End Property
Public Property Get getOBSERVACIONES_DET() As String
    getOBSERVACIONES_DET = OBSERVACIONES_DET
End Property
Public Property Get getES_DUPLICADO() As Integer
    getES_DUPLICADO = ES_DUPLICADO
End Property
Public Function InsertarGranulometria() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID_GRANULOMETRIA
    consulta = "Insert into granulometrias " & _
               " values(" & _
               ID_GRANULOMETRIA & "," & MUESTRA_ID & "," & DETERMINACION_ID & ",'" & _
               PESO_100 & "','" & PARTIDOS & "','" & MERMADOS & "','" & OTROS_CEREALES & "','" & _
               TRIGO_BLANDO & "','" & ATACADOS & "','" & CON_COLORACION & "','" & MACULADOS & "','" & _
               FUSARIADOS & "','" & GERMINADOS & "','" & SEM_NOCIVAS & "','" & SEM_NO_NOCIVAS & "','" & _
               AVERIADOS & "','" & OTROS_DANADOS & "','" & PMD & "','" & INSECTOS & "','" & _
               VITREOS & "','" & BERRENDOS & "','" & HARINOSOS & "'," & _
               ES_DUPLICADO & ",'" & _
               PESO_MUESTRA & "','" & _
               OBSERVACIONES_DET & "')"
    execute_bd consulta
    InsertarGranulometria = ID_GRANULOMETRIA
    Exit Function
fallo:
    InsertarGranulometria = 0
    MsgBox "Error al insertar las Granulometrias (InsertarGranulometria)", vbCritical, Err.Description
End Function

Public Sub CrearID_GRANULOMETRIA()
    Dim rs As New ADOdb.RecordSet
    Dim consulta As String
    consulta = "SELECT MAX(ID_GRANULOMETRIA) FROM GRANULOMETRIAS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_GRANULOMETRIA = 1
    Else
        ID_GRANULOMETRIA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function CargarGranulometria(ID As Long) As Boolean
    Dim consulta As String
    Dim rs As New ADOdb.RecordSet
    consulta = "SELECT * FROM granulometrias WHERE ID_GRANULOMETRIA=" & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
         CargarGranulometria = False
    Else
        ID_GRANULOMETRIA = rs.Fields("ID_GRANULOMETRIA")
        MUESTRA_ID = rs.Fields("muestra_id")
        DETERMINACION_ID = rs.Fields("DETERMINACION_ID")
        If Not IsNull(rs.Fields("PESO_100")) Then
            PESO_100 = rs.Fields("PESO_100")
        Else
            PESO_100 = 0
        End If
        If Not IsNull(rs.Fields("PARTIDOS")) Then
            PARTIDOS = rs.Fields("PARTIDOS")
        Else
            PARTIDOS = 0
        End If
        If Not IsNull(rs.Fields("MERMADOS")) Then
            MERMADOS = rs.Fields("MERMADOS")
        Else
            MERMADOS = 0
        End If
        If Not IsNull(rs.Fields("OTROS_CEREALES")) Then
            OTROS_CEREALES = rs.Fields("OTROS_CEREALES")
        Else
            OTROS_CEREALES = 0
        End If
        If Not IsNull(rs.Fields("TRIGO_BLANDO")) Then
            TRIGO_BLANDO = rs.Fields("TRIGO_BLANDO")
        Else
            TRIGO_BLANDO = 0
        End If
        If Not IsNull(rs.Fields("ATACADOS")) Then
            ATACADOS = rs.Fields("ATACADOS")
        Else
            ATACADOS = 0
        End If
        If Not IsNull(rs.Fields("CON_COLORACION")) Then
            CON_COLORACION = rs.Fields("CON_COLORACION")
        Else
            CON_COLORACION = 0
        End If
        If Not IsNull(rs.Fields("MACULADOS")) Then
            MACULADOS = rs.Fields("MACULADOS")
        Else
            MACULADOS = 0
        End If
        If Not IsNull(rs.Fields("FUSARIADOS")) Then
            FUSARIADOS = rs.Fields("FUSARIADOS")
        Else
            FUSARIADOS = 0
        End If
        If Not IsNull(rs.Fields("GERMINADOS")) Then
            GERMINADOS = rs.Fields("GERMINADOS")
        Else
            GERMINADOS = 0
        End If
        If Not IsNull(rs.Fields("SEM_NOCIVAS")) Then
            SEM_NOCIVAS = rs.Fields("SEM_NOCIVAS")
        Else
            SEM_NOCIVAS = 0
        End If
        If Not IsNull(rs.Fields("SEM_NO_NOCIVAS")) Then
            SEM_NO_NOCIVAS = rs.Fields("SEM_NO_NOCIVAS")
        Else
            SEM_NO_NOCIVAS = 0
        End If
        If Not IsNull(rs.Fields("AVERIADOS")) Then
            AVERIADOS = rs.Fields("AVERIADOS")
        Else
            AVERIADOS = 0
        End If
        If Not IsNull(rs.Fields("OTROS_DANADOS")) Then
            OTROS_DANADOS = rs.Fields("OTROS_DANADOS")
        Else
            OTROS_DANADOS = 0
        End If
        If Not IsNull(rs.Fields("PMD")) Then
            PMD = rs.Fields("PMD")
        Else
            PMD = 0
        End If
        If Not IsNull(rs.Fields("INSECTOS")) Then
            INSECTOS = rs.Fields("INSECTOS")
        Else
            INSECTOS = 0
        End If
        If Not IsNull(rs.Fields("VITREOS")) Then
            VITREOS = rs.Fields("VITREOS")
        Else
            VITREOS = 0
        End If
        If Not IsNull(rs.Fields("BERRENDOS")) Then
            BERRENDOS = rs.Fields("BERRENDOS")
        Else
            BERRENDOS = 0
        End If
        If Not IsNull(rs.Fields("HARINOSOS")) Then
            HARINOSOS = rs.Fields("HARINOSOS")
        Else
            HARINOSOS = 0
        End If
        If Not IsNull(rs.Fields("PESO_MUESTRA")) Then
            PESO_MUESTRA = rs.Fields("PESO_MUESTRA")
        Else
            PESO_MUESTRA = 0
        End If
        If Not IsNull(rs.Fields("OBSERVACIONES_DET")) Then
            OBSERVACIONES_DET = rs.Fields("OBSERVACIONES_DET")
        Else
            OBSERVACIONES_DET = 0
        End If
        ES_DUPLICADO = rs.Fields("ES_DUPLICADO")
    End If
    Set rs = Nothing
    CargarGranulometria = True
    Exit Function
fallo:
    CargarGranulometria = False
    MsgBox "Error al cargar la Granulometria (CargarGranulometria)", vbCritical, Err.Description
End Function
Public Function ComprobarGranulometria(MUESTRA As Long, deter As Long) As Long
    Dim consulta As String
    Dim rs As New ADOdb.RecordSet
    consulta = "SELECT id_granulometria FROM granulometrias WHERE MUESTRA_ID=" & MUESTRA & " AND DETERMINACION_ID=" & deter
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ComprobarGranulometria = 0
    Else
        ComprobarGranulometria = rs("id_granulometria")
    End If
    Set rs = Nothing
    Exit Function
fallo:
    ComprobarGranulometria = 0
    MsgBox "Error al comprobar la Granulometria (ComprobarGranulometria)", vbCritical, Err.Description
End Function

Public Function Modificar(granulometria As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "update granulometrias set " & _
               " PESO_100 = '" & PESO_100 & "'," & _
               " PARTIDOS = '" & PARTIDOS & "'," & _
               " MERMADOS = '" & MERMADOS & "'," & _
               " OTROS_CEREALES = '" & OTROS_CEREALES & "'," & _
               " TRIGO_BLANDO = '" & TRIGO_BLANDO & "'," & _
               " ATACADOS = '" & ATACADOS & "'," & _
               " CON_COLORACION = '" & CON_COLORACION & "'," & _
               " MACULADOS = '" & MACULADOS & "'," & _
               " FUSARIADOS = '" & FUSARIADOS & "'," & _
               " GERMINADOS = '" & GERMINADOS & "'," & _
               " SEM_NOCIVAS = '" & SEM_NOCIVAS & "'," & _
               " SEM_NO_NOCIVAS = '" & SEM_NO_NOCIVAS & "'," & _
               " AVERIADOS = '" & AVERIADOS & "'," & _
               " OTROS_DANADOS = '" & OTROS_DANADOS & "'," & _
               " PMD = '" & PMD & "'," & " INSECTOS = '" & INSECTOS & "'," & _
               " VITREOS = '" & VITREOS & "'," & _
               " BERRENDOS = '" & BERRENDOS & "'," & " HARINOSOS = '" & HARINOSOS & "'," & _
               " ES_DUPLICADO = " & ES_DUPLICADO & "," & _
               " PESO_MUESTRA = '" & PESO_MUESTRA & "'," & _
               " OBSERVACIONES_DET = '" & OBSERVACIONES_DET & "'" & _
               " WHERE ID_GRANULOMETRIA = " & granulometria
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar la Granulometria (Modificar)", vbCritical, Err.Description
End Function


