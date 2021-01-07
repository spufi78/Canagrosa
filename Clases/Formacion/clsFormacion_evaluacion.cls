VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_evaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_EVALUACION
'***************************************************************
Private CURSO_ID As Long
Private EMPLEADO_ID As Long
Private ORGANIZACION As Long
Private nivel As Long
Private UTILIDAD As Long
Private CASOS_PRACTICOS As Long
Private AUDIOVISUALES As Long
Private DINAMICAS As Long
Private COMODIDAD As Long
Private DURACION As Long
Private HORARIO As Long
Private MATERIAL As Long
Private GENERAL As Long
Private OBSERVACIONES As String
Private FECHA_EVALUACION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_EVALUACION
'***************************************************************
Public Property Let setCURSO_ID(ByVal dato As Long)
        CURSO_ID = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setORGANIZACION(ByVal dato As Long)
        ORGANIZACION = dato
End Property
Public Property Let setNIVEL(ByVal dato As Long)
        nivel = dato
End Property
Public Property Let setUTILIDAD(ByVal dato As Long)
        UTILIDAD = dato
End Property
Public Property Let setCASOS_PRACTICOS(ByVal dato As Long)
        CASOS_PRACTICOS = dato
End Property
Public Property Let setAUDIOVISUALES(ByVal dato As Long)
        AUDIOVISUALES = dato
End Property
Public Property Let setDINAMICAS(ByVal dato As Long)
        DINAMICAS = dato
End Property
Public Property Let setCOMODIDAD(ByVal dato As Long)
        COMODIDAD = dato
End Property
Public Property Let setDURACION(ByVal dato As Long)
        DURACION = dato
End Property
Public Property Let setHORARIO(ByVal dato As Long)
        HORARIO = dato
End Property
Public Property Let setMATERIAL(ByVal dato As Long)
        MATERIAL = dato
End Property
Public Property Let setGENERAL(ByVal dato As Long)
        GENERAL = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setFECHA_EVALUACION(ByVal dato As String)
        FECHA_EVALUACION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_EVALUACION
'***************************************************************
Public Property Get getCURSO_ID() As Long
        getCURSO_ID = CURSO_ID
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getORGANIZACION() As Long
        getORGANIZACION = ORGANIZACION
End Property
Public Property Get getNIVEL() As Long
        getNIVEL = nivel
End Property
Public Property Get getUTILIDAD() As Long
        getUTILIDAD = UTILIDAD
End Property
Public Property Get getCASOS_PRACTICOS() As Long
        getCASOS_PRACTICOS = CASOS_PRACTICOS
End Property
Public Property Get getAUDIOVISUALES() As Long
        getAUDIOVISUALES = AUDIOVISUALES
End Property
Public Property Get getDINAMICAS() As Long
        getDINAMICAS = DINAMICAS
End Property
Public Property Get getCOMODIDAD() As Long
        getCOMODIDAD = COMODIDAD
End Property
Public Property Get getDURACION() As Long
        getDURACION = DURACION
End Property
Public Property Get getHORARIO() As Long
        getHORARIO = HORARIO
End Property
Public Property Get getMATERIAL() As Long
        getMATERIAL = MATERIAL
End Property
Public Property Get getGENERAL() As Long
        getGENERAL = GENERAL
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFECHA_EVALUACION() As String
        getFECHA_EVALUACION = FECHA_EVALUACION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_EVALUACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_EVALUACION WHERE CURSO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CURSO_ID = rs("CURSO_ID")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                ORGANIZACION = rs("ORGANIZACION")
                nivel = rs("NIVEL")
                UTILIDAD = rs("UTILIDAD")
                CASOS_PRACTICOS = rs("CASOS_PRACTICOS")
                AUDIOVISUALES = rs("AUDIOVISUALES")
                DINAMICAS = rs("DINAMICAS")
                COMODIDAD = rs("COMODIDAD")
                DURACION = rs("DURACION")
                HORARIO = rs("HORARIO")
                MATERIAL = rs("MATERIAL")
                GENERAL = rs("GENERAL")
                OBSERVACIONES = rs("OBSERVACIONES")
                FECHA_EVALUACION = rs("FECHA_EVALUACION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_evaluacion)", vbCritical, Err.Description
End Function

Public Function Carga_usuario(ID As Long, ID_USUARIO As Long) As Boolean
        On Error GoTo fallo
        
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * FROM FORMACION_EVALUACION WHERE CURSO_ID = " & ID & " AND EMPLEADO_ID =" & ID_USUARIO
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_usuario = False
        Else
                CURSO_ID = rs("CURSO_ID")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                ORGANIZACION = rs("ORGANIZACION")
                nivel = rs("NIVEL")
                UTILIDAD = rs("UTILIDAD")
                CASOS_PRACTICOS = rs("CASOS_PRACTICOS")
                AUDIOVISUALES = rs("AUDIOVISUALES")
                DINAMICAS = rs("DINAMICAS")
                COMODIDAD = rs("COMODIDAD")
                DURACION = rs("DURACION")
                HORARIO = rs("HORARIO")
                MATERIAL = rs("MATERIAL")
                GENERAL = rs("GENERAL")
                OBSERVACIONES = rs("OBSERVACIONES")
                FECHA_EVALUACION = rs("FECHA_EVALUACION")
                Carga_usuario = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_usuario = False
        MsgBox "Error al cargar los datos(clsFormacion_evaluacion)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FORMACION_EVALUACION " & _
                   "  VALUES (" & _
                        CURSO_ID & "," & EMPLEADO_ID & "," & ORGANIZACION & "," & _
                        nivel & "," & UTILIDAD & "," & CASOS_PRACTICOS & "," & _
                        AUDIOVISUALES & "," & DINAMICAS & "," & COMODIDAD & "," & _
                        DURACION & "," & HORARIO & "," & MATERIAL & "," & _
                        GENERAL & "," & "'" & OBSERVACIONES & "',CURRENT_TIMESTAMP)"
        execute_bd consulta
        Insertar = CURSO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_evaluacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_EVALUACION SET " & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " ORGANIZACION = " & ORGANIZACION & "," & _
                        " NIVEL = " & nivel & "," & _
                        " UTILIDAD = " & UTILIDAD & "," & _
                        " CASOS_PRACTICOS = " & CASOS_PRACTICOS & "," & _
                        " AUDIOVISUALES = " & AUDIOVISUALES & "," & _
                        " DINAMICAS = " & DINAMICAS & "," & _
                        " COMODIDAD = " & COMODIDAD & "," & _
                        " DURACION = " & DURACION & "," & _
                        " HORARIO = " & HORARIO & "," & _
                        " MATERIAL = " & MATERIAL & "," & _
                        " GENERAL = " & GENERAL & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                        " FECHA_EVALUACION = '" & FECHA_EVALUACION & "'" & _
                " WHERE CURSO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_evaluacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_EVALUACION " & _
                "    WHERE CURSO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_evaluacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_EVALUACION ORDER BY CURSO_ID"
        Set Listado = datos_bd(consulta)
End Function
