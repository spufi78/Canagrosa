VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFluidos_ip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFLUIDOS_IP
'***************************************************************
Private MUESTRA_ID As Long
Private TIPO_EQUIPO As Long
Private BOTE_EX_ID As Long
Private NUMERO_PARTICULAS As Long
Private LIMPIEZA_CONFORME As Long
Private LIMPIEZA_OB As String
Private INSPECCION_CONFORME As Long
Private VOLUMEN As String
Private TIPO_CALIBRACION As Long
Private PRIMER_CONTEO As Long
Private CONTADOR_ID As Long
Private FECHA_ULT_CALIBRACION As String
Private SENSOR_ID As Long
Private FLUJO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFLUIDOS_IP
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTIPO_EQUIPO(ByVal dato As Long)
        TIPO_EQUIPO = dato
End Property
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setNUMERO_PARTICULAS(ByVal dato As Long)
        NUMERO_PARTICULAS = dato
End Property
Public Property Let setLIMPIEZA_CONFORME(ByVal dato As Long)
        LIMPIEZA_CONFORME = dato
End Property
Public Property Let setLIMPIEZA_OB(ByVal dato As String)
        LIMPIEZA_OB = dato
End Property
Public Property Let setINSPECCION_CONFORME(ByVal dato As Long)
        INSPECCION_CONFORME = dato
End Property
Public Property Let setVOLUMEN(ByVal dato As String)
        VOLUMEN = dato
End Property
Public Property Let setTIPO_CALIBRACION(ByVal dato As Long)
        TIPO_CALIBRACION = dato
End Property
Public Property Let setPRIMER_CONTEO(ByVal dato As Long)
        PRIMER_CONTEO = dato
End Property
Public Property Let setCONTADOR_ID(ByVal dato As Long)
        CONTADOR_ID = dato
End Property
Public Property Let setFECHA_ULT_CALIBRACION(ByVal dato As String)
        FECHA_ULT_CALIBRACION = dato
End Property
Public Property Let setSENSOR_ID(ByVal dato As Long)
        SENSOR_ID = dato
End Property
Public Property Let setFLUJO(ByVal dato As String)
        FLUJO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFLUIDOS_IP
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO_EQUIPO() As Long
        getTIPO_EQUIPO = TIPO_EQUIPO
End Property
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getNUMERO_PARTICULAS() As Long
        getNUMERO_PARTICULAS = NUMERO_PARTICULAS
End Property
Public Property Get getLIMPIEZA_CONFORME() As Long
        getLIMPIEZA_CONFORME = LIMPIEZA_CONFORME
End Property
Public Property Get getLIMPIEZA_OB() As String
        getLIMPIEZA_OB = LIMPIEZA_OB
End Property
Public Property Get getINSPECCION_CONFORME() As Long
        getINSPECCION_CONFORME = INSPECCION_CONFORME
End Property
Public Property Get getVOLUMEN() As String
        getVOLUMEN = VOLUMEN
End Property
Public Property Get getTIPO_CALIBRACION() As Long
        getTIPO_CALIBRACION = TIPO_CALIBRACION
End Property
Public Property Get getPRIMER_CONTEO() As Long
        getPRIMER_CONTEO = PRIMER_CONTEO
End Property
Public Property Get getCONTADOR_ID() As Long
        getCONTADOR_ID = CONTADOR_ID
End Property
Public Property Get getFECHA_ULT_CALIBRACION() As String
        getFECHA_ULT_CALIBRACION = FECHA_ULT_CALIBRACION
End Property
Public Property Get getSENSOR_ID() As Long
        getSENSOR_ID = SENSOR_ID
End Property
Public Property Get getFLUJO() As String
        getFLUJO = FLUJO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFLUIDOS_IP
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM FLUIDOS_IP WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                TIPO_EQUIPO = rs("TIPO_EQUIPO")
                BOTE_EX_ID = rs("BOTE_EX_ID")
                NUMERO_PARTICULAS = rs("NUMERO_PARTICULAS")
                LIMPIEZA_CONFORME = rs("LIMPIEZA_CONFORME")
                LIMPIEZA_OB = rs("LIMPIEZA_OB")
                INSPECCION_CONFORME = rs("INSPECCION_CONFORME")
                VOLUMEN = rs("VOLUMEN")
                TIPO_CALIBRACION = rs("TIPO_CALIBRACION")
                PRIMER_CONTEO = rs("PRIMER_CONTEO")
                
                If IsNull(rs("CONTADOR_ID")) Then
                    CONTADOR_ID = 0
                Else
                    CONTADOR_ID = rs("CONTADOR_ID")
                End If
                If IsNull(rs("FECHA_ULT_CALIBRACION")) Then
                    FECHA_ULT_CALIBRACION = ""
                Else
                    FECHA_ULT_CALIBRACION = rs("FECHA_ULT_CALIBRACION")
                End If
                If IsNull(rs("SENSOR_ID")) Then
                    SENSOR_ID = 0
                Else
                    SENSOR_ID = rs("SENSOR_ID")
                End If
                If IsNull(rs("FLUJO")) Then
                    FLUJO = ""
                Else
                    FLUJO = rs("FLUJO")
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFluidos_ip)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FLUIDOS_IP " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & TIPO_EQUIPO & "," & BOTE_EX_ID & "," & _
                        NUMERO_PARTICULAS & "," & LIMPIEZA_CONFORME & "," & "'" & LIMPIEZA_OB & "'" & "," & _
                        INSPECCION_CONFORME & "," & "'" & VOLUMEN & "'" & "," & TIPO_CALIBRACION & "," & _
                        PRIMER_CONTEO & "," & CONTADOR_ID & ",'" & FECHA_ULT_CALIBRACION & "'," & SENSOR_ID & ",'" & FLUJO & "'" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFluidos_ip)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FLUIDOS_IP SET " & _
                        " TIPO_EQUIPO = " & TIPO_EQUIPO & "," & _
                        " BOTE_EX_ID = " & BOTE_EX_ID & "," & _
                        " NUMERO_PARTICULAS = " & NUMERO_PARTICULAS & "," & _
                        " LIMPIEZA_CONFORME = " & LIMPIEZA_CONFORME & "," & _
                        " LIMPIEZA_OB = '" & LIMPIEZA_OB & "'," & _
                        " INSPECCION_CONFORME = " & INSPECCION_CONFORME & "," & _
                        " VOLUMEN = '" & VOLUMEN & "'," & _
                        " TIPO_CALIBRACION = " & TIPO_CALIBRACION & "," & _
                        " PRIMER_CONTEO = " & PRIMER_CONTEO & "," & _
                        " CONTADOR_ID = " & CONTADOR_ID & "," & _
                        " FECHA_ULT_CALIBRACION = '" & FECHA_ULT_CALIBRACION & "'," & _
                        " SENSOR_ID = " & SENSOR_ID & "," & _
                        " FLUJO = '" & FLUJO & "'" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFluidos_ip)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FLUIDOS_IP " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFluidos_ip)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM FLUIDOS_IP ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function
