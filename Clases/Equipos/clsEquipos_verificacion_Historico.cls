VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_verificacion_hco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'************************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_VERIFICACION_HCO
'************************************************************************
Private VERIFICACION_ID As Long
Private FECHA As String
Private MODALIDAD As String
Private PROCEDIMIENTO As String
Private OPERADOR_ID As Long
Private OPERADOR As String
Private LIMITACIONES_USO As String
Private RANGO_MIN As String
Private RANGO_MAX As String
Private UNIDADES As String
Private CONFORME As Long

'**************************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_VERIFICACION_HCO
'**************************************************************************
Public Property Let setVERIFICACION_ID(ByVal dato As Long)
        VERIFICACION_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        FECHA = dato
End Property
Public Property Let setMODALIDAD(ByVal dato As String)
        MODALIDAD = dato
End Property
Public Property Let setPROCEDIMIENTO(ByVal dato As String)
        PROCEDIMIENTO = dato
End Property
Public Property Let setOPERADOR_ID(ByVal dato As Long)
        OPERADOR_ID = dato
End Property
Public Property Let setOPERADOR(ByVal dato As String)
        OPERADOR = dato
End Property
Public Property Let setLIMITACIONES_USO(ByVal dato As String)
        LIMITACIONES_USO = dato
End Property
Public Property Let setRANGO_MIN(ByVal dato As String)
        RANGO_MIN = dato
End Property
Public Property Let setRANGO_MAX(ByVal dato As String)
        RANGO_MAX = dato
End Property
Public Property Let setUNIDADES(ByVal dato As String)
        UNIDADES = dato
End Property
Public Property Let setCONFORME(ByVal dato As Long)
        CONFORME = dato
End Property

'************************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_VERIFICACION_HCO
'************************************************************************
Public Property Get getVERIFICACION_ID() As Long
        getVERIFICACION_ID = VERIFICACION_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = FECHA
End Property
Public Property Get getMODALIDAD() As String
        getMODALIDAD = MODALIDAD
End Property
Public Property Get getPROCEDIMIENTO() As String
        getPROCEDIMIENTO = PROCEDIMIENTO
End Property
Public Property Get getOPERADOR_ID() As Long
        getOPERADOR_ID = OPERADOR_ID
End Property
Public Property Get getOPERADOR() As String
        getOPERADOR = OPERADOR
End Property
Public Property Get getLIMITACIONES_USO() As String
        getLIMITACIONES_USO = LIMITACIONES_USO
End Property
Public Property Get getRANGO_MIN() As String
        getRANGO_MIN = RANGO_MIN
End Property
Public Property Get getRANGO_MAX() As String
        getRANGO_MAX = RANGO_MAX
End Property
Public Property Get getUNIDADES() As String
        getUNIDADES = UNIDADES
End Property
Public Property Get getCONFORME() As Long
        getCONFORME = CONFORME
End Property

'****************************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_VERIFICACION_HCO
'****************************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_VERIFICACION_HISTORICO WHERE VERIFICACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            VERIFICACION_ID = rs("VERIFICACION_ID")
            FECHA = rs("FECHA")
            MODALIDAD = rs("MODALIDAD")
            PROCEDIMIENTO = rs("PROCEDIMIENTO")
            OPERADOR_ID = rs("OPERADOR_ID")
            OPERADOR = rs("OPERADOR")
            LIMITACIONES_USO = rs("LIMITACIONES_USO")
            RANGO_MIN = rs("RANGO_MIN")
            RANGO_MAX = rs("RANGO_MAX")
            UNIDADES = rs("UNIDADES")
            CONFORME = rs("CONFORME")
            
            Carga = True
        End If
        
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_verificacion_hco)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "INSERT INTO EQUIPOS_VERIFICACION_HISTORICO " & _
                   "VALUES (" & _
                        VERIFICACION_ID & "," & _
                        "'" & Format(FECHA, "yyyy-mm-dd hh:nn:ss") & "'" & "," & "'" & MODALIDAD & "'" & "," & _
                        "'" & PROCEDIMIENTO & "'" & ", " & OPERADOR_ID & ", " & "'" & OPERADOR & "'" & "," & _
                        "'" & LIMITACIONES_USO & "'" & ", " & _
                        "'" & RANGO_MIN & "'" & ", " & "'" & RANGO_MAX & "'" & ", " & _
                        "'" & UNIDADES & "'" & ", " & CONFORME & ", " & _
                        "'" & usuario.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                   ")"
        
        execute_bd consulta
        Insertar = VERIFICACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_verificacion_hco)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long, strFecha As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_VERIFICACION_HISTORICO " & _
                   "   SET MODALIDAD = '" & MODALIDAD & "'" & _
                   "     , PROCEDIMIENTO = '" & PROCEDIMIENTO & "'" & _
                   "     , OPERADOR_ID = " & OPERADOR_ID & _
                   "     , OPERADOR = '" & OPERADOR & "'" & _
                   "     , LIMITACIONES_USO = '" & LIMITACIONES_USO & "'" & _
                   "     , RANGO_MIN = '" & RANGO_MIN & "'" & _
                   "     , RANGO_MAX = '" & RANGO_MAX & "'" & _
                   "     , UNIDADES = '" & UNIDADES & "'" & _
                   "     , CONFORME = " & CONFORME & _
                   "     , CUSERID = '" & usuario.getUSUARIO & "'" & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE VERIFICACION_ID = " & ID & _
                   "   AND FECHA = '" & Format(strFecha, "yyyy-mm-dd hh:nn:ss") & "'"
        
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos_verificacion_hco)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long, strFecha As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM EQUIPOS_VERIFICACION_HISTORICO " & _
                   " WHERE VERIFICACION_ID = " & ID & _
                   "   AND FECHA = '" & strFecha & "'"
                   
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_verificacion_hco)", vbCritical, Err.Description
End Function

' listado de historico de verificaciones del equipo y verificación pasados por parámetro
Public Function Listado(ID_V As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT EVH.VERIFICACION_ID, EVH.FECHA, EVH.FECHA, EVH.MODALIDAD, EVH.OPERADOR " & _
                   "     , EVH.LIMITACIONES_USO, EVH.RANGO_MIN, EVH.RANGO_MAX, EVH.UNIDADES " & _
                   "     , IF(EVH.CONFORME='1','SI', 'NO') AS CONFORME " & _
                   "     , EVH.OPERADOR_ID " & _
                   "  FROM EQUIPOS_VERIFICACION_HISTORICO EVH " & _
                   " WHERE EVH.VERIFICACION_ID = " & ID_V & _
                   " ORDER BY EVH.FECHA DESC "
        
        Set Listado = datos_bd(consulta)
End Function
