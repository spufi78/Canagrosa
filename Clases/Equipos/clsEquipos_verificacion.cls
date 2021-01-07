VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_verificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_VERIFICACION
'**************************************************************
Private ID_VERIFICACION As Long
Private EQUIPO_ID As Long
Private MODALIDAD_ID As Long
Private PERIODICIDAD_ID As Long
Private PROCEDIMIENTO_ID As Long
Private VERIFICADOR_EXTERNO_ID As Long
Private VERIFICADOR_INTERNO_ID As Long
Private fecha_actual As String
Private fecha_proxima As String
Private RANGO_MIN As String
Private RANGO_MAX As String
Private UNIDADES_ID As Long
Private ACTIVA As Long
Private EFECTIVA As Long

Private ID_AUX As Integer
Public Property Let setID_AUX(ByVal dato As Integer)
    ID_AUX = dato
End Property
Public Property Get getID_AUX() As Integer
    getID_AUX = ID_AUX
End Property

'****************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_VERIFICACION
'****************************************************************
Public Property Let setID_VERIFICACION(ByVal dato As Long)
        ID_VERIFICACION = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setMODALIDAD_ID(ByVal dato As Long)
        MODALIDAD_ID = dato
End Property
Public Property Let setPERIODICIDAD_ID(ByVal dato As Long)
        PERIODICIDAD_ID = dato
End Property
Public Property Let setPROCEDIMIENTO_ID(ByVal dato As Long)
        PROCEDIMIENTO_ID = dato
End Property
Public Property Let setVERIFICADOR_EXTERNO_ID(ByVal dato As Long)
        VERIFICADOR_EXTERNO_ID = dato
End Property
Public Property Let setVERIFICADOR_INTERNO_ID(ByVal dato As Long)
        VERIFICADOR_INTERNO_ID = dato
End Property
Public Property Let setFECHA_ACTUAL(ByVal dato As String)
        fecha_actual = dato
End Property
Public Property Let setFECHA_PROXIMA(ByVal dato As String)
        fecha_proxima = dato
End Property
Public Property Let setRANGO_MIN(ByVal dato As String)
        RANGO_MIN = dato
End Property
Public Property Let setRANGO_MAX(ByVal dato As String)
        RANGO_MAX = dato
End Property
Public Property Let setUNIDADES_ID(ByVal dato As Long)
        UNIDADES_ID = dato
End Property
Public Property Let setACTIVA(ByVal dato As Long)
        ACTIVA = dato
End Property
Public Property Let setEFECTIVA(ByVal dato As Long)
        EFECTIVA = dato
End Property

'**************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_VERIFICACION
'**************************************************************
Public Property Get getID_VERIFICACION() As Long
        getID_VERIFICACION = ID_VERIFICACION
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getMODALIDAD_ID() As Long
        getMODALIDAD_ID = MODALIDAD_ID
End Property
Public Property Get getPERIODICIDAD_ID() As Long
        getPERIODICIDAD_ID = PERIODICIDAD_ID
End Property
Public Property Get getPROCEDIMIENTO_ID() As Long
        getPROCEDIMIENTO_ID = PROCEDIMIENTO_ID
End Property
Public Property Get getVERIFICADOR_EXTERNO_ID() As Long
        getVERIFICADOR_EXTERNO_ID = VERIFICADOR_EXTERNO_ID
End Property
Public Property Get getVERIFICADOR_INTERNO_ID() As Long
        getVERIFICADOR_INTERNO_ID = VERIFICADOR_INTERNO_ID
End Property
Public Property Get getFECHA_ACTUAL() As String
        getFECHA_ACTUAL = fecha_actual
End Property
Public Property Get getFECHA_PROXIMA() As String
        getFECHA_PROXIMA = fecha_proxima
End Property
Public Property Get getRANGO_MIN() As String
        getRANGO_MIN = RANGO_MIN
End Property
Public Property Get getRANGO_MAX() As String
        getRANGO_MAX = RANGO_MAX
End Property
Public Property Get getUNIDADES_ID() As Long
        getUNIDADES_ID = UNIDADES_ID
End Property
Public Property Get getACTIVA() As Long
        getACTIVA = ACTIVA
End Property
Public Property Get getEFECTIVA() As Long
        getEFECTIVA = EFECTIVA
End Property

'******************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_VERIFICACION
'******************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_VERIFICACION WHERE ID_VERIFICACION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            ID_VERIFICACION = rs("ID_VERIFICACION")
            EQUIPO_ID = rs("EQUIPO_ID")
            MODALIDAD_ID = rs("MODALIDAD_ID")
            PERIODICIDAD_ID = rs("PERIODICIDAD_ID")
            PROCEDIMIENTO_ID = rs("PROCEDIMIENTO_ID")
            VERIFICADOR_EXTERNO_ID = rs("VERIFICADOR_EXTERNO_ID")
            VERIFICADOR_INTERNO_ID = rs("VERIFICADOR_INTERNO_ID")
            fecha_actual = Format(rs("FECHA_ACTUAL"), "yyyy-mm-dd")
            fecha_proxima = Format(rs("FECHA_PROXIMA"), "yyyy-mm-dd")
            RANGO_MIN = rs("RANGO_MIN")
            RANGO_MAX = rs("RANGO_MAX")
            UNIDADES_ID = rs("UNIDADES_ID")
            ACTIVA = rs("ACTIVA")
            EFECTIVA = rs("EFECTIVA")
            
            Carga = True
        End If
        
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_verificacion)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        Me.CrearId_Verificacion
        
        consulta = "INSERT INTO EQUIPOS_VERIFICACION " & _
                   "VALUES (" & _
                        ID_VERIFICACION & "," & EQUIPO_ID & "," & _
                        MODALIDAD_ID & "," & PERIODICIDAD_ID & "," & _
                        PROCEDIMIENTO_ID & "," & VERIFICADOR_EXTERNO_ID & ", " & _
                        VERIFICADOR_INTERNO_ID & ", " & "'" & Format(fecha_actual, "yyyy-mm-dd") & "'" & ", " & _
                        "'" & Format(fecha_proxima, "yyyy-mm-dd") & "'" & ", " & "'" & RANGO_MIN & "'" & ", " & _
                        "'" & RANGO_MAX & "'" & ", " & UNIDADES_ID & ", " & ACTIVA & ", " & EFECTIVA & ", " & _
                        "'" & usuario.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                   ")"
        
        execute_bd consulta
        Insertar = ID_VERIFICACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_verificacion)", vbCritical, Err.Description
End Function

Public Sub CrearId_Verificacion()
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT MAX(ID_VERIFICACION) " & _
               "  FROM EQUIPOS_VERIFICACION "
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_VERIFICACION = 1
    Else
        ID_VERIFICACION = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_VERIFICACION " & _
                   "   SET MODALIDAD_ID = " & MODALIDAD_ID & _
                   "     , PERIODICIDAD_ID = " & PERIODICIDAD_ID & _
                   "     , PROCEDIMIENTO_ID = " & PROCEDIMIENTO_ID & _
                   "     , VERIFICADOR_EXTERNO_ID = " & VERIFICADOR_EXTERNO_ID & _
                   "     , VERIFICADOR_INTERNO_ID = " & VERIFICADOR_INTERNO_ID & _
                   "     , FECHA_ACTUAL = '" & Format(fecha_actual, "yyyy-mm-dd") & "'" & _
                   "     , FECHA_PROXIMA = '" & Format(fecha_proxima, "yyyy-mm-dd") & "'" & _
                   "     , RANGO_MIN = '" & RANGO_MIN & "'" & _
                   "     , RANGO_MAX = '" & RANGO_MAX & "'" & _
                   "     , UNIDADES_ID = " & UNIDADES_ID & _
                   "     , ACTIVA = " & ACTIVA & _
                   "     , EFECTIVA = " & EFECTIVA & _
                   "     , CUSERID = '" & usuario.getUSUARIO & "'" & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE ID_VERIFICACION = " & ID

        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos_verificacion)", vbCritical, Err.Description
End Function

' se establece ACTIVA = 0
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_VERIFICACION " & _
                   "   SET ACTIVA = 0 " & _
                   " WHERE ID_VERIFICACION = " & ID
        
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_verificacion)", vbCritical, Err.Description
End Function

' se borran todas las verificaciones del equipo (ACTIVA = 0)
Public Function eliminar_todas_verificaciones(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_VERIFICACION " & _
                   "   SET ACTIVA = 0 " & _
                   " WHERE EQUIPO_ID = " & ID
        
        execute_bd consulta
        eliminar_todas_verificaciones = True
        Exit Function
fallo:
        eliminar_todas_verificaciones = False
        MsgBox "Error al eliminar_todas_verificaciones (clsEquipos_verificacion)", vbCritical, Err.Description
End Function

Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * " & _
                   "  FROM EQUIPOS_VERIFICACION " & _
                   " WHERE ACTIVA = 1 " & _
                   "   AND EFECTIVA = 1 " & _
                   " ORDER BY ID_VERIFICACION "
        
        Set Listado = datos_bd(consulta)
End Function

Public Function Listado_verificaciones_asignadas(lngEquipo As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ID_VERIFICACION " & _
                   "     , if(isnull(MODALIDAD.DESCRIPCION),'', MODALIDAD.DESCRIPCION) AS MODALIDAD " & _
                   "     , if(isnull(PERIODICIDAD.DESCRIPCION),'', PERIODICIDAD.DESCRIPCION) AS PERIODICIDAD " & _
                   "     , CONCAT(CONCAT(CONCAT(USUARIOS.NOMBRE, ' '), USUARIOS.APELLIDOS), PROVEEDORES.NOMBRE) AS RESPONSABLE " & _
                   "  FROM EQUIPOS_VERIFICACION EV " & _
                   "     , DECODIFICADORA AS MODALIDAD " & _
                   "     , DECODIFICADORA AS PERIODICIDAD " & _
                   "     , USUARIOS, PROVEEDORES " & _
                   " WHERE EV.MODALIDAD_ID = MODALIDAD.VALOR " & _
                   "   AND EV.PERIODICIDAD_ID = PERIODICIDAD.VALOR " & _
                   "   AND EV.VERIFICADOR_INTERNO_ID = USUARIOS.ID_EMPLEADO " & _
                   "   AND EV.VERIFICADOR_EXTERNO_ID = PROVEEDORES.ID_PROVEEDOR " & _
                   "   AND MODALIDAD.CODIGO = " & decodificadora.EQ_TIPO_CALIBRACION & _
                   "   AND PERIODICIDAD.CODIGO = " & decodificadora.EQ_periodicidad & _
                   "   AND EQUIPO_ID = " & lngEquipo & _
                   "   AND ACTIVA = 1 " & _
                   " ORDER BY ID_VERIFICACION"

        
        Set Listado_verificaciones_asignadas = datos_bd(consulta)
End Function

Public Function total_verificaciones(lngEquipo As Long) As Long
        Dim rs As New ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT COUNT(*) AS TOTAL " & _
                   "  FROM EQUIPOS_VERIFICACION EV" & _
                   " WHERE EV.EQUIPO_ID = " & lngEquipo & _
                   "   AND EV.ACTIVA = 1 "

        Set rs = datos_bd(consulta)
        total_verificaciones = rs("TOTAL")
        
        Set rs = Nothing
End Function

'Public Function total_verificaciones_con_nombre(lngEquipo As Long, strNombre As String) As Long
'        Dim rs As New ADODB.RecordSet
'        Dim consulta As String
'
'        consulta = "SELECT COUNT(*) AS TOTAL " & _
'                   "  FROM EQUIPOS_VERIFICACION EV" & _
'                   " WHERE EV.EQUIPO_ID = " & lngEquipo & _
'                   "   AND EV.NOMBRE = '" & strNombre & "' " & _
'                   "   AND EV.ACTIVA = 1 "
'
'        Set rs = datos_bd(consulta)
'        total_verificaciones_con_nombre = rs("TOTAL")
'
'        Set rs = Nothing
'End Function

Public Function Verificacion_mas_cercana(lngEquipo As Long) As Long
        Dim rs As New ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ID_VERIFICACION " & _
                   "  FROM EQUIPOS_VERIFICACION EV" & _
                   " WHERE EV.EQUIPO_ID = " & lngEquipo & _
                   "   AND EV.ACTIVA = 1 " & _
                   " ORDER BY FECHA_PROXIMA "
        
        Set rs = datos_bd(consulta)
        
        Verificacion_mas_cercana = rs(0)
        
        Set rs = Nothing
End Function

' función que marca como efectiva la calibración del equipo pasado por parámetro
Public Function marcar_no_efectivas_como_efectivas(lngEquipo As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_VERIFICACION " & _
                   "   SET EFECTIVA = 1 " & _
                   "     , CUSERID = '" & usuario.getUSUARIO & "'" & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE EQUIPO_ID = " & lngEquipo & _
                   "   AND EFECTIVA = 0 "
        
        execute_bd consulta
        marcar_no_efectivas_como_efectivas = True
        Exit Function
fallo:
        marcar_no_efectivas_como_efectivas = False
        MsgBox "Error al marcar no efectivas como efectivas (clsEquipos_verificacion)", vbCritical, Err.Description
End Function

Public Function Eliminar_no_efectivas(lngEquipo As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM EQUIPOS_VERIFICACION " & _
                   " WHERE EFECTIVA = 0 " & _
                   "   AND EQUIPO_ID = " & lngEquipo
                   
        execute_bd consulta
        Eliminar_no_efectivas = True
        Exit Function
fallo:
        Eliminar_no_efectivas = False
        MsgBox "Error al Eliminar no efectivas (clsEquipos_verificacion)", vbCritical, Err.Description
End Function
