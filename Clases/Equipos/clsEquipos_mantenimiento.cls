VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_mantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_MANTENIMIENTO
'***************************************************************
Private EQUIPO_ID As Long
Private DIARIO_MODALIDAD_ID As Long
Private DIARIO_RESPONSABLE_INTERNO_ID As Long
Private DIARIO_RESPONSABLE_EXTERNO_ID As Long
Private DIARIO_MANTENIMIENTO As Integer
Private DIARIO_PROCEDIMIENTO As String
Private SEMANAL_MODALIDAD_ID As Long
Private SEMANAL_RESPONSABLE_INTERNO_ID As Long
Private SEMANAL_RESPONSABLE_EXTERNO_ID As Long
Private SEMANAL_MANTENIMIENTO As Integer
Private SEMANAL_REGISTRO As String
Private SEMANAL_PROCEDIMIENTO As String
Private SEMANAL_FECHA As String
Private MENSUAL_MODALIDAD_ID As Long
Private MENSUAL_RESPONSABLE_INTERNO_ID As Long
Private MENSUAL_RESPONSABLE_EXTERNO_ID As Long
Private MENSUAL_MANTENIMIENTO As Integer
Private MENSUAL_REGISTRO As String
Private MENSUAL_PROCEDIMIENTO As String
Private MENSUAL_FECHA As String
Private MENSUAL_ENERO As String
Private MENSUAL_FEBRERO As String
Private MENSUAL_MARZO As String
Private MENSUAL_ABRIL As String
Private MENSUAL_MAYO As String
Private MENSUAL_JUNIO As String
Private MENSUAL_JULIO As String
Private MENSUAL_AGOSTO As String
Private MENSUAL_SEPTIEMBRE As String
Private MENSUAL_OCTUBRE As String
Private MENSUAL_NOVIEMBRE As String
Private MENSUAL_DICIEMBRE As String
Private BIMESTRAL_MANTENIMIENTO As Integer
Private SEMESTRAL_MANTENIMIENTO As Integer
Private ANUAL_MANTENIMIENTO As Integer
Private TRIMESTRAL_MANTENIMIENTO As Integer

Private ID_AUX As Integer
Public Property Let setID_AUX(ByVal dato As Integer)
    ID_AUX = dato
End Property
Public Property Get getID_AUX() As Integer
    getID_AUX = ID_AUX
End Property

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_MANTENIMIENTO
'***************************************************************
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setDIARIO_MODALIDAD_ID(ByVal dato As Long)
        DIARIO_MODALIDAD_ID = dato
End Property
Public Property Let setDIARIO_RESPONSABLE_INTERNO_ID(ByVal dato As Long)
        DIARIO_RESPONSABLE_INTERNO_ID = dato
End Property
Public Property Let setDIARIO_RESPONSABLE_EXTERNO_ID(ByVal dato As Long)
        DIARIO_RESPONSABLE_EXTERNO_ID = dato
End Property
Public Property Let setDIARIO_MANTENIMIENTO(ByVal dato As Integer)
        DIARIO_MANTENIMIENTO = dato
End Property
Public Property Let setDIARIO_PROCEDIMIENTO(ByVal dato As String)
        DIARIO_PROCEDIMIENTO = dato
End Property
Public Property Let setSEMANAL_MODALIDAD_ID(ByVal dato As Long)
        SEMANAL_MODALIDAD_ID = dato
End Property
Public Property Let setSEMANAL_RESPONSABLE_INTERNO_ID(ByVal dato As Long)
        SEMANAL_RESPONSABLE_INTERNO_ID = dato
End Property
Public Property Let setSEMANAL_RESPONSABLE_EXTERNO_ID(ByVal dato As Long)
        SEMANAL_RESPONSABLE_EXTERNO_ID = dato
End Property
Public Property Let setSEMANAL_MANTENIMIENTO(ByVal dato As Integer)
        SEMANAL_MANTENIMIENTO = dato
End Property
Public Property Let setSEMANAL_REGISTRO(ByVal dato As String)
        SEMANAL_REGISTRO = dato
End Property
Public Property Let setSEMANAL_PROCEDIMIENTO(ByVal dato As String)
        SEMANAL_PROCEDIMIENTO = dato
End Property
Public Property Let setSEMANAL_FECHA(ByVal dato As String)
        SEMANAL_FECHA = dato
End Property
Public Property Let setMENSUAL_MODALIDAD_ID(ByVal dato As Long)
        MENSUAL_MODALIDAD_ID = dato
End Property
Public Property Let setMENSUAL_RESPONSABLE_INTERNO_ID(ByVal dato As Long)
        MENSUAL_RESPONSABLE_INTERNO_ID = dato
End Property
Public Property Let setMENSUAL_RESPONSABLE_EXTERNO_ID(ByVal dato As Long)
        MENSUAL_RESPONSABLE_EXTERNO_ID = dato
End Property
Public Property Let setMENSUAL_MANTENIMIENTO(ByVal dato As Integer)
        MENSUAL_MANTENIMIENTO = dato
End Property
Public Property Let setMENSUAL_REGISTRO(ByVal dato As String)
        MENSUAL_REGISTRO = dato
End Property
Public Property Let setMENSUAL_PROCEDIMIENTO(ByVal dato As String)
        MENSUAL_PROCEDIMIENTO = dato
End Property
Public Property Let setMENSUAL_FECHA(ByVal dato As String)
        MENSUAL_FECHA = dato
End Property
Public Property Let setMENSUAL_ENERO(ByVal dato As String)
        MENSUAL_ENERO = dato
End Property
Public Property Let setMENSUAL_FEBRERO(ByVal dato As String)
        MENSUAL_FEBRERO = dato
End Property
Public Property Let setMENSUAL_MARZO(ByVal dato As String)
        MENSUAL_MARZO = dato
End Property
Public Property Let setMENSUAL_ABRIL(ByVal dato As String)
        MENSUAL_ABRIL = dato
End Property
Public Property Let setMENSUAL_MAYO(ByVal dato As String)
        MENSUAL_MAYO = dato
End Property
Public Property Let setMENSUAL_JUNIO(ByVal dato As String)
        MENSUAL_JUNIO = dato
End Property
Public Property Let setMENSUAL_JULIO(ByVal dato As String)
        MENSUAL_JULIO = dato
End Property
Public Property Let setMENSUAL_AGOSTO(ByVal dato As String)
        MENSUAL_AGOSTO = dato
End Property
Public Property Let setMENSUAL_SEPTIEMBRE(ByVal dato As String)
        MENSUAL_SEPTIEMBRE = dato
End Property
Public Property Let setMENSUAL_OCTUBRE(ByVal dato As String)
        MENSUAL_OCTUBRE = dato
End Property
Public Property Let setMENSUAL_NOVIEMBRE(ByVal dato As String)
        MENSUAL_NOVIEMBRE = dato
End Property
Public Property Let setMENSUAL_DICIEMBRE(ByVal dato As String)
        MENSUAL_DICIEMBRE = dato
End Property
Public Property Let setBIMESTRAL_MANTENIMIENTO(ByVal dato As Integer)
        BIMESTRAL_MANTENIMIENTO = dato
End Property
Public Property Let setSEMESTRAL_MANTENIMIENTO(ByVal dato As Integer)
        SEMESTRAL_MANTENIMIENTO = dato
End Property
Public Property Let setANUAL_MANTENIMIENTO(ByVal dato As Integer)
        ANUAL_MANTENIMIENTO = dato
End Property
Public Property Let setTRIMESTRAL_MANTENIMIENTO(ByVal dato As Integer)
        TRIMESTRAL_MANTENIMIENTO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_MANTENIMIENTO
'***************************************************************
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getDIARIO_MODALIDAD_ID() As Long
        getDIARIO_MODALIDAD_ID = DIARIO_MODALIDAD_ID
End Property
Public Property Get getDIARIO_RESPONSABLE_INTERNO_ID() As Long
        getDIARIO_RESPONSABLE_INTERNO_ID = DIARIO_RESPONSABLE_INTERNO_ID
End Property
Public Property Get getDIARIO_RESPONSABLE_EXTERNO_ID() As Long
        getDIARIO_RESPONSABLE_EXTERNO_ID = DIARIO_RESPONSABLE_EXTERNO_ID
End Property
Public Property Get getDIARIO_MANTENIMIENTO() As Integer
        getDIARIO_MANTENIMIENTO = DIARIO_MANTENIMIENTO
End Property
Public Property Get getDIARIO_PROCEDIMIENTO() As String
        getDIARIO_PROCEDIMIENTO = DIARIO_PROCEDIMIENTO
End Property
Public Property Get getSEMANAL_MODALIDAD_ID() As Long
        getSEMANAL_MODALIDAD_ID = SEMANAL_MODALIDAD_ID
End Property
Public Property Get getSEMANAL_RESPONSABLE_INTERNO_ID() As Long
        getSEMANAL_RESPONSABLE_INTERNO_ID = SEMANAL_RESPONSABLE_INTERNO_ID
End Property
Public Property Get getSEMANAL_RESPONSABLE_EXTERNO_ID() As Long
        getSEMANAL_RESPONSABLE_EXTERNO_ID = SEMANAL_RESPONSABLE_EXTERNO_ID
End Property
Public Property Get getSEMANAL_MANTENIMIENTO() As Integer
        getSEMANAL_MANTENIMIENTO = SEMANAL_MANTENIMIENTO
End Property
Public Property Get getSEMANAL_REGISTRO() As String
        getSEMANAL_REGISTRO = SEMANAL_REGISTRO
End Property
Public Property Get getSEMANAL_PROCEDIMIENTO() As String
        getSEMANAL_PROCEDIMIENTO = SEMANAL_PROCEDIMIENTO
End Property
Public Property Get getSEMANAL_FECHA() As String
        getSEMANAL_FECHA = SEMANAL_FECHA
End Property
Public Property Get getMENSUAL_MODALIDAD_ID() As Long
        getMENSUAL_MODALIDAD_ID = MENSUAL_MODALIDAD_ID
End Property
Public Property Get getMENSUAL_RESPONSABLE_INTERNO_ID() As Long
        getMENSUAL_RESPONSABLE_INTERNO_ID = MENSUAL_RESPONSABLE_INTERNO_ID
End Property
Public Property Get getMENSUAL_RESPONSABLE_EXTERNO_ID() As Long
        getMENSUAL_RESPONSABLE_EXTERNO_ID = MENSUAL_RESPONSABLE_EXTERNO_ID
End Property
Public Property Get getMENSUAL_MANTENIMIENTO() As Integer
        getMENSUAL_MANTENIMIENTO = MENSUAL_MANTENIMIENTO
End Property
Public Property Get getMENSUAL_REGISTRO() As String
        getMENSUAL_REGISTRO = MENSUAL_REGISTRO
End Property
Public Property Get getMENSUAL_PROCEDIMIENTO() As String
        getMENSUAL_PROCEDIMIENTO = MENSUAL_PROCEDIMIENTO
End Property
Public Property Get getMENSUAL_FECHA() As String
        getMENSUAL_FECHA = MENSUAL_FECHA
End Property
Public Property Get getMENSUAL_ENERO() As String
        getMENSUAL_ENERO = MENSUAL_ENERO
End Property
Public Property Get getMENSUAL_FEBRERO() As String
        getMENSUAL_FEBRERO = MENSUAL_FEBRERO
End Property
Public Property Get getMENSUAL_MARZO() As String
        getMENSUAL_MARZO = MENSUAL_MARZO
End Property
Public Property Get getMENSUAL_ABRIL() As String
        getMENSUAL_ABRIL = MENSUAL_ABRIL
End Property
Public Property Get getMENSUAL_MAYO() As String
        getMENSUAL_MAYO = MENSUAL_MAYO
End Property
Public Property Get getMENSUAL_JUNIO() As String
        getMENSUAL_JUNIO = MENSUAL_JUNIO
End Property
Public Property Get getMENSUAL_JULIO() As String
        getMENSUAL_JULIO = MENSUAL_JULIO
End Property
Public Property Get getMENSUAL_AGOSTO() As String
        getMENSUAL_AGOSTO = MENSUAL_AGOSTO
End Property
Public Property Get getMENSUAL_SEPTIEMBRE() As String
        getMENSUAL_SEPTIEMBRE = MENSUAL_SEPTIEMBRE
End Property
Public Property Get getMENSUAL_OCTUBRE() As String
        getMENSUAL_OCTUBRE = MENSUAL_OCTUBRE
End Property
Public Property Get getMENSUAL_NOVIEMBRE() As String
        getMENSUAL_NOVIEMBRE = MENSUAL_NOVIEMBRE
End Property
Public Property Get getMENSUAL_DICIEMBRE() As String
        getMENSUAL_DICIEMBRE = MENSUAL_DICIEMBRE
End Property
Public Property Get getBIMESTRAL_MANTENIMIENTO() As Integer
        getBIMESTRAL_MANTENIMIENTO = BIMESTRAL_MANTENIMIENTO
End Property
Public Property Get getSEMESTRAL_MANTENIMIENTO() As Integer
        getSEMESTRAL_MANTENIMIENTO = SEMESTRAL_MANTENIMIENTO
End Property
Public Property Get getANUAL_MANTENIMIENTO() As Integer
        getANUAL_MANTENIMIENTO = ANUAL_MANTENIMIENTO
End Property
Public Property Get getTRIMESTRAL_MANTENIMIENTO() As Integer
        getTRIMESTRAL_MANTENIMIENTO = TRIMESTRAL_MANTENIMIENTO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_MANTENIMIENTO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM EQUIPOS_MANTENIMIENTO WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                EQUIPO_ID = rs("EQUIPO_ID")
                DIARIO_MODALIDAD_ID = rs("DIARIO_MODALIDAD_ID")
                DIARIO_RESPONSABLE_INTERNO_ID = rs("DIARIO_RESPONSABLE_INTERNO_ID")
                DIARIO_RESPONSABLE_EXTERNO_ID = rs("DIARIO_RESPONSABLE_EXTERNO_ID")
                DIARIO_MANTENIMIENTO = rs("DIARIO_MANTENIMIENTO")
                DIARIO_PROCEDIMIENTO = rs("DIARIO_PROCEDIMIENTO")
                SEMANAL_MODALIDAD_ID = rs("SEMANAL_MODALIDAD_ID")
                SEMANAL_RESPONSABLE_INTERNO_ID = rs("SEMANAL_RESPONSABLE_INTERNO_ID")
                SEMANAL_RESPONSABLE_EXTERNO_ID = rs("SEMANAL_RESPONSABLE_EXTERNO_ID")
                SEMANAL_MANTENIMIENTO = rs("SEMANAL_MANTENIMIENTO")
                SEMANAL_REGISTRO = rs("SEMANAL_REGISTRO")
                SEMANAL_PROCEDIMIENTO = rs("SEMANAL_PROCEDIMIENTO")
                SEMANAL_FECHA = rs("SEMANAL_FECHA")
                MENSUAL_MODALIDAD_ID = rs("MENSUAL_MODALIDAD_ID")
                MENSUAL_RESPONSABLE_INTERNO_ID = rs("MENSUAL_RESPONSABLE_INTERNO_ID")
                MENSUAL_RESPONSABLE_EXTERNO_ID = rs("MENSUAL_RESPONSABLE_EXTERNO_ID")
                MENSUAL_MANTENIMIENTO = rs("MENSUAL_MANTENIMIENTO")
                MENSUAL_REGISTRO = rs("MENSUAL_REGISTRO")
                MENSUAL_PROCEDIMIENTO = rs("MENSUAL_PROCEDIMIENTO")
                MENSUAL_FECHA = rs("MENSUAL_FECHA")
                MENSUAL_ENERO = rs("MENSUAL_ENERO")
                MENSUAL_FEBRERO = rs("MENSUAL_FEBRERO")
                MENSUAL_MARZO = rs("MENSUAL_MARZO")
                MENSUAL_ABRIL = rs("MENSUAL_ABRIL")
                MENSUAL_MAYO = rs("MENSUAL_MAYO")
                MENSUAL_JUNIO = rs("MENSUAL_JUNIO")
                MENSUAL_JULIO = rs("MENSUAL_JULIO")
                MENSUAL_AGOSTO = rs("MENSUAL_AGOSTO")
                MENSUAL_SEPTIEMBRE = rs("MENSUAL_SEPTIEMBRE")
                MENSUAL_OCTUBRE = rs("MENSUAL_OCTUBRE")
                MENSUAL_NOVIEMBRE = rs("MENSUAL_NOVIEMBRE")
                MENSUAL_DICIEMBRE = rs("MENSUAL_DICIEMBRE")
                BIMESTRAL_MANTENIMIENTO = rs("BIMESTRAL_MANTENIMIENTO")
                SEMESTRAL_MANTENIMIENTO = rs("SEMESTRAL_MANTENIMIENTO")
                ANUAL_MANTENIMIENTO = rs("ANUAL_MANTENIMIENTO")
                TRIMESTRAL_MANTENIMIENTO = rs("TRIMESTRAL_MANTENIMIENTO")
                
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_mantenimiento)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "INSERT INTO EQUIPOS_MANTENIMIENTO " & _
                   "  VALUES (" & _
                        EQUIPO_ID & "," & _
                        DIARIO_MODALIDAD_ID & ", " & DIARIO_RESPONSABLE_INTERNO_ID & ", " & DIARIO_RESPONSABLE_EXTERNO_ID & ", " & _
                        DIARIO_MANTENIMIENTO & "," & "'" & DIARIO_PROCEDIMIENTO & "'" & "," & _
                        SEMANAL_MODALIDAD_ID & ", " & SEMANAL_RESPONSABLE_INTERNO_ID & ", " & SEMANAL_RESPONSABLE_EXTERNO_ID & ", " & _
                        SEMANAL_MANTENIMIENTO & "," & "'" & SEMANAL_REGISTRO & "'" & "," & _
                        "'" & SEMANAL_PROCEDIMIENTO & "'" & "," & "'" & SEMANAL_FECHA & "'" & "," & _
                        MENSUAL_MODALIDAD_ID & ", " & MENSUAL_RESPONSABLE_INTERNO_ID & ", " & MENSUAL_RESPONSABLE_EXTERNO_ID & ", " & _
                        MENSUAL_MANTENIMIENTO & "," & "'" & MENSUAL_REGISTRO & "'" & "," & "'" & MENSUAL_PROCEDIMIENTO & "'" & "," & _
                        "'" & MENSUAL_FECHA & "'" & "," & "'" & MENSUAL_ENERO & "'" & "," & _
                        "'" & MENSUAL_FEBRERO & "'" & "," & "'" & MENSUAL_MARZO & "'" & "," & "'" & MENSUAL_ABRIL & "'" & "," & _
                        "'" & MENSUAL_MAYO & "'" & "," & "'" & MENSUAL_JUNIO & "'" & "," & "'" & MENSUAL_JULIO & "'" & "," & _
                        "'" & MENSUAL_AGOSTO & "'" & "," & "'" & MENSUAL_SEPTIEMBRE & "'" & "," & "'" & MENSUAL_OCTUBRE & "'" & "," & _
                        "'" & MENSUAL_NOVIEMBRE & "'" & "," & "'" & MENSUAL_DICIEMBRE & "'" & "," & BIMESTRAL_MANTENIMIENTO & "," & _
                        SEMESTRAL_MANTENIMIENTO & "," & ANUAL_MANTENIMIENTO & "," & TRIMESTRAL_MANTENIMIENTO & ", " & _
                        "'" & usuario.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                   ")"
        
        execute_bd consulta
        Insertar = EQUIPO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_mantenimiento)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    
    'consulta = "UPDATE EQUIPOS_MANTENIMIENTO SET " & _
            " DIARIO_MANTENIMIENTO = " & DIARIO_MANTENIMIENTO & "," & " DIARIO_PROCEDIMIENTO = '" & DIARIO_PROCEDIMIENTO & "'," & _
            " DIARIO_RESPONSABLE = '" & DIARIO_RESPONSABLE & "'," & " SEMANAL_MANTENIMIENTO = " & SEMANAL_MANTENIMIENTO & "," & _
            " SEMANAL_REGISTRO = '" & SEMANAL_REGISTRO & "'," & " SEMANAL_PROCEDIMIENTO = '" & SEMANAL_PROCEDIMIENTO & "'," & _
            " SEMANAL_FECHA = '" & SEMANAL_FECHA & "'," & " SEMANAL_RESPONSABLE = '" & SEMANAL_RESPONSABLE & "'," & _
            " MENSUAL_MANTENIMIENTO = " & MENSUAL_MANTENIMIENTO & "," & " MENSUAL_REGISTRO = '" & MENSUAL_REGISTRO & "'," & _
            " MENSUAL_PROCEDIMIENTO = '" & MENSUAL_PROCEDIMIENTO & "'," & " MENSUAL_FECHA = '" & MENSUAL_FECHA & "'," & _
            " MENSUAL_RESPONSABLE = '" & MENSUAL_RESPONSABLE & "'," & " MENSUAL_ENERO = '" & MENSUAL_ENERO & "'," & _
            " MENSUAL_FEBRERO = '" & MENSUAL_FEBRERO & "'," & " MENSUAL_MARZO = '" & MENSUAL_MARZO & "'," & _
            " MENSUAL_ABRIL = '" & MENSUAL_ABRIL & "'," & " MENSUAL_MAYO = '" & MENSUAL_MAYO & "'," & _
            " MENSUAL_JUNIO = '" & MENSUAL_JUNIO & "'," & " MENSUAL_JULIO = '" & MENSUAL_JULIO & "'," & _
            " MENSUAL_AGOSTO = '" & MENSUAL_AGOSTO & "'," & " MENSUAL_SEPTIEMBRE = '" & MENSUAL_SEPTIEMBRE & "'," & _
            " MENSUAL_OCTUBRE = '" & MENSUAL_OCTUBRE & "'," & " MENSUAL_NOVIEMBRE = '" & MENSUAL_NOVIEMBRE & "'," & _
            " BIMESTRAL_MANTENIMIENTO = " & BIMESTRAL_MANTENIMIENTO & "," & " SEMESTRAL_MANTENIMIENTO = " & SEMESTRAL_MANTENIMIENTO & "," & _
            " ANUAL_MANTENIMIENTO = " & ANUAL_MANTENIMIENTO & _
            " MENSUAL_DICIEMBRE = '" & MENSUAL_DICIEMBRE & "'," & " TRIMESTRAL_MANTENIMIENTO = " & TRIMESTRAL_MANTENIMIENTO & ", " & _
            " CUSERID = '" & usuario.getUSUARIO & "', " & _
            " FTIMESTP = CURRENT_TIMESTAMP " & _
        " WHERE EQUIPO_ID = " & ID

    consulta = "UPDATE EQUIPOS_MANTENIMIENTO SET " & _
               "       DIARIO_MODALIDAD_ID = " & DIARIO_MODALIDAD_ID & _
               "     , DIARIO_RESPONSABLE_INTERNO_ID = " & DIARIO_RESPONSABLE_INTERNO_ID & _
               "     , DIARIO_RESPONSABLE_EXTERNO_ID = " & DIARIO_RESPONSABLE_EXTERNO_ID & _
               "     , DIARIO_MANTENIMIENTO = " & DIARIO_MANTENIMIENTO & _
               "     , DIARIO_PROCEDIMIENTO = '" & DIARIO_PROCEDIMIENTO & "'" & _
               "     , SEMANAL_MODALIDAD_ID = " & SEMANAL_MODALIDAD_ID & _
               "     , SEMANAL_RESPONSABLE_INTERNO_ID = " & SEMANAL_RESPONSABLE_INTERNO_ID & _
               "     , SEMANAL_RESPONSABLE_EXTERNO_ID = " & SEMANAL_RESPONSABLE_EXTERNO_ID & _
               "     , SEMANAL_MANTENIMIENTO = " & SEMANAL_MANTENIMIENTO & _
               "     , SEMANAL_REGISTRO = '" & SEMANAL_REGISTRO & "'" & _
               "     , SEMANAL_PROCEDIMIENTO = '" & SEMANAL_PROCEDIMIENTO & "'"
    consulta = consulta & _
               "     , SEMANAL_FECHA = '" & SEMANAL_FECHA & "'" & _
               "     , MENSUAL_MODALIDAD_ID = " & MENSUAL_MODALIDAD_ID & _
               "     , MENSUAL_RESPONSABLE_INTERNO_ID = " & MENSUAL_RESPONSABLE_INTERNO_ID & _
               "     , MENSUAL_RESPONSABLE_EXTERNO_ID = " & MENSUAL_RESPONSABLE_EXTERNO_ID & _
               "     , MENSUAL_MANTENIMIENTO = " & MENSUAL_MANTENIMIENTO & _
               "     , MENSUAL_REGISTRO = '" & MENSUAL_REGISTRO & "'" & _
               "     , MENSUAL_PROCEDIMIENTO = '" & MENSUAL_PROCEDIMIENTO & "'" & _
               "     , MENSUAL_FECHA = '" & MENSUAL_FECHA & "'"
    consulta = consulta & _
               "     , MENSUAL_ENERO = '" & MENSUAL_ENERO & "'" & _
               "     , MENSUAL_FEBRERO = '" & MENSUAL_FEBRERO & "'" & _
               "     , MENSUAL_MARZO = '" & MENSUAL_MARZO & "'" & _
               "     , MENSUAL_ABRIL = '" & MENSUAL_ABRIL & "'" & _
               "     , MENSUAL_MAYO = '" & MENSUAL_MAYO & "'" & _
               "     , MENSUAL_JUNIO = '" & MENSUAL_JUNIO & "'" & _
               "     , MENSUAL_JULIO = '" & MENSUAL_JULIO & "'" & _
               "     , MENSUAL_AGOSTO = '" & MENSUAL_AGOSTO & "'" & _
               "     , MENSUAL_SEPTIEMBRE = '" & MENSUAL_SEPTIEMBRE & "'" & _
               "     , MENSUAL_OCTUBRE = '" & MENSUAL_OCTUBRE & "'" & _
               "     , MENSUAL_NOVIEMBRE = '" & MENSUAL_NOVIEMBRE & "'" & _
               "     , BIMESTRAL_MANTENIMIENTO = " & BIMESTRAL_MANTENIMIENTO & _
               "     , SEMESTRAL_MANTENIMIENTO = " & SEMESTRAL_MANTENIMIENTO & _
               "     , ANUAL_MANTENIMIENTO = " & ANUAL_MANTENIMIENTO & _
               "     , MENSUAL_DICIEMBRE = '" & MENSUAL_DICIEMBRE & "'" & _
               "     , TRIMESTRAL_MANTENIMIENTO = " & TRIMESTRAL_MANTENIMIENTO & _
               "     , CUSERID = '" & usuario.getUSUARIO & "'" & _
               "     , FTIMESTP = CURRENT_TIMESTAMP " & _
               " WHERE EQUIPO_ID = " & ID
    
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar (clsEquipos_mantenimiento)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "DELETE FROM EQUIPOS_MANTENIMIENTO " & _
               " WHERE EQUIPO_ID = " & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al Eliminar (clsEquipos_mantenimiento)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT * FROM EQUIPOS_MANTENIMIENTO ORDER BY EQUIPO_ID"
    Set Listado = datos_bd(consulta)
End Function
