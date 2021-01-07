VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_planes_Mantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_PLANES_MANTENIMIENTO
'***************************************************************
Private ID_PLAN_MTO As Long
Private NOMBRE As String
Private DESCRIPCION As String
Private FRECUENCIA_ID As Long
Private PLANIF_MESES As String
Private PLANIF_DIAS_SEMANA As String
Private PLANIF_DIAS_MES As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_PLANES_MANTENIMIENTO
'***************************************************************
Public Property Let setID_PLAN_MTO(ByVal dato As Long)
        ID_PLAN_MTO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFRECUENCIA_ID(ByVal dato As Long)
        FRECUENCIA_ID = dato
End Property
Public Property Let setPLANIF_MESES(ByVal dato As String)
        PLANIF_MESES = dato
End Property
Public Property Let setPLANIF_DIAS_SEMANA(ByVal dato As String)
        PLANIF_DIAS_SEMANA = dato
End Property
Public Property Let setPLANIF_DIAS_MES(ByVal dato As String)
        PLANIF_DIAS_MES = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_PLANES_MANTENIMIENTO
'***************************************************************
Public Property Get getID_PLAN_MTO() As Long
        getID_PLAN_MTO = ID_PLAN_MTO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFRECUENCIA_ID() As Long
        getFRECUENCIA_ID = FRECUENCIA_ID
End Property
Public Property Get getPLANIF_MESES() As String
        getPLANIF_MESES = PLANIF_MESES
End Property
Public Property Get getPLANIF_DIAS_SEMANA() As String
        getPLANIF_DIAS_SEMANA = PLANIF_DIAS_SEMANA
End Property
Public Property Get getPLANIF_DIAS_MES() As String
        getPLANIF_DIAS_MES = PLANIF_DIAS_MES
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_PLANES_MANTENIMIENTO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM eq_planes_mantenimiento WHERE ID_PLAN_MTO = " & id
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PLAN_MTO = rs("ID_PLAN_MTO")
                NOMBRE = rs("NOMBRE")
                DESCRIPCION = rs("DESCRIPCION")
                FRECUENCIA_ID = rs("FRECUENCIA_ID")
                PLANIF_MESES = rs("PLANIF_MESES")
                PLANIF_DIAS_SEMANA = rs("PLANIF_DIAS_SEMANA")
                PLANIF_DIAS_MES = rs("PLANIF_DIAS_MES")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        log ("Error al cargar los datos(clsEQUIPOS_PLANES_MANTENIMIENTO)")
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_PLAN_MTO) FROM eq_planes_mantenimiento"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PLAN_MTO = 1
        Else
                ID_PLAN_MTO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
        
End Function
Public Function Insertar() As String
    On Error GoTo fallo
    Dim consulta As String
    
    Me.CrearID
    If Not duplicado(ID_PLAN_MTO) Then
        consulta = "INSERT INTO eq_planes_mantenimiento " & _
                   "  VALUES (" & _
                        ID_PLAN_MTO & "," & "'" & NOMBRE & "'" & "," & "'" & DESCRIPCION & "'" & ", " & _
                        FRECUENCIA_ID & ", " & "'" & PLANIF_MESES & "'" & ", " & "'" & PLANIF_DIAS_SEMANA & "'" & ", " & "'" & PLANIF_DIAS_MES & "'" & ", " & _
                        ORDEN & ", " & _
                        "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                   ")"
        execute_bd consulta
        Insertar = ID_PLAN_MTO
    Else
        Insertar = 0
    End If
    Exit Function
    
fallo:
    Insertar = 0
    MsgBox "Error al insertar (clsEQUIPOS_PLANES_MANTENIMIENTO)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    
    Modificar = False
    If Not duplicado(ID) Then
        consulta = "UPDATE eq_planes_mantenimiento SET " & _
                   "       NOMBRE = '" & NOMBRE & "'" & ", " & _
                   "       DESCRIPCION = '" & DESCRIPCION & "'" & ", " & _
                   "       FRECUENCIA_ID = " & FRECUENCIA_ID & ", " & _
                   "       PLANIF_MESES = '" & PLANIF_MESES & "', " & _
                   "       PLANIF_DIAS_SEMANA = '" & PLANIF_DIAS_SEMANA & "', " & _
                   "       PLANIF_DIAS_MES = '" & PLANIF_DIAS_MES & "', " & _
                   "       CUSERID = '" & USUARIO.getUSUARIO & "', " & _
                   "       FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE ID_PLAN_MTO = " & ID
        
        execute_bd consulta
        Modificar = True
    End If
    
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar (clsEQUIPOS_PLANES_MANTENIMIENTO)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.RecordSet

    ' Se comprueba si hay equipos con ese plan asignado
    consulta = "SELECT COUNT(*) AS TOTAL " & _
               "  FROM EQUIPOS_PLANES_MTO_ASIGNADOS " & _
               " WHERE PLAN_MTO_ID = " & ID
    Set rs = datos_bd(consulta)
    If rs("TOTAL") <> 0 Then
        MsgBox "No se puede eliminar el plan de mantenimiento. Está asociado a algún equipo.", vbExclamation, App.Title
        Eliminar = False
        Exit Function
    End If
    ' Se borra el detalle y el plan
    execute_bd ("DELETE FROM eq_planes_mantenimiento_DETALLE WHERE PLAN_MTO_ID = " & id)
    execute_bd ("DELETE FROM eq_planes_mantenimiento WHERE ID_PLAN_MTO = " & id)
    Eliminar = True
    
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al Eliminar (clsEQUIPOS_PLANES_MANTENIMIENTO)", vbCritical, Err.Description
End Function
Public Function Listado(NOMBRE As String, DESCRIPCION As String, FRECUENCIA As String) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT EQPM.ID_PLAN_MTO, EQPM.NOMBRE Plan, EQPM.DESCRIPCION Descripcion, FREQ.DESCRIPCION Frecuencia " & _
                   "  FROM eq_planes_mantenimiento EQPM, DECODIFICADORA FREQ " & _
                   " WHERE FREQ.CODIGO = " & decodificadora.EQ_periodicidad & _
                   "   AND EQPM.FRECUENCIA_ID = FREQ.VALOR " & _
                   "   AND EQPM.NOMBRE LIKE '%" & NOMBRE & "%'" & _
                   "   AND EQPM.DESCRIPCION LIKE '%" & DESCRIPCION & "%'" & _
                   "   AND FREQ.VALOR LIKE '%" & FRECUENCIA & "%'" & _
                   " ORDER BY EQPM.ORDEN"
                   
        Set Listado = datos_bd(consulta)
End Function
Public Function Frecuencia_descrip(ID_FREQ As Long) As String
    Dim consulta As String
    Dim rs As New ADODB.RecordSet
    
    consulta = "SELECT FREQ.DESCRIPCION Frecuencia " & _
               "  FROM DECODIFICADORA FREQ " & _
               " WHERE FREQ.CODIGO = " & decodificadora.EQ_periodicidad & _
               "   AND FREQ.VALOR = " & ID_FREQ
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        MsgBox "No se encontró la frecuencia.", vbExclamation, App.Title
        Frecuencia_descrip = 0
        Exit Function
    Else
        Frecuencia_descrip = rs.Fields("Frecuencia")
    End If

End Function
Public Function Listado_completo() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM eq_planes_mantenimiento" & _
                   " ORDER BY ORDEN"
        Set Listado_completo = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ID_PLAN_MTO, NOMBRE FROM eq_planes_mantenimiento ORDER BY ORDEN"
        Set Listado_Combo = datos_bd(consulta)
End Function

Public Function duplicado(lngID_PLAN As Long) As Boolean
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT ID_PLAN_MTO " & _
               "  FROM eq_planes_mantenimiento " & _
               " WHERE NOMBRE = '" & NOMBRE & "'" & _
               "   AND ID_PLAN_MTO <> " & lngID_PLAN
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recuperó ninguno
        duplicado = False
    Else
        MsgBox "Existe un plan con el mismo nombre.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
