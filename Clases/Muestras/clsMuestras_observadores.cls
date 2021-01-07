VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestras_observadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMUESTRAS_OBSERVADORES
'***************************************************************
Private MUESTRA_ID As Long
Private DETERMINACION_ENSAYO_ID As Long
Private TIPO_DETERMINACION_ENSAYO_ID As Long
'M0961-I
Private SELLANTE_ID As Long
Private ENSAYO As Long
'M0961-F
Private DOCUMENTO_ID As Long
Private fecha As String
Private ES_CONTROL_EFICACIA As Integer
Private USUARIO_ID_REALIZACION As Long
Private USUARIO_ID_OBSERVADOR As Long
Private USUARIO_RESULTADO1 As Long
Private USUARIO_RESULTADO2 As Long
Private REALIZACION_ES_FORMADOR As Integer
Private OBSERVADOR_ES_FORMADOR As Integer
Private ES_DUPLICADO As Integer
Private ES_RECUALIFICACION As Integer
Private OBSERVACIONES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMUESTRAS_OBSERVADORES
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setDETERMINACION_ENSAYO_ID(ByVal dato As Long)
        DETERMINACION_ENSAYO_ID = dato
End Property
Public Property Let setTIPO_DETERMINACION_ENSAYO_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ENSAYO_ID = dato
End Property
'M0961-I
Public Property Let setSELLANTE_ID(ByVal dato As Long)
        SELLANTE_ID = dato
End Property
Public Property Let setENSAYO(ByVal dato As Long)
        ENSAYO = dato
End Property
'M0961-F
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setES_CONTROL_EFICACIA(ByVal dato As Integer)
        ES_CONTROL_EFICACIA = dato
End Property
Public Property Let setUSUARIO_ID_REALIZACION(ByVal dato As Long)
        USUARIO_ID_REALIZACION = dato
End Property
Public Property Let setUSUARIO_ID_OBSERVADOR(ByVal dato As Long)
        USUARIO_ID_OBSERVADOR = dato
End Property
Public Property Let setUSUARIO_RESULTADO1(ByVal dato As Long)
        USUARIO_RESULTADO1 = dato
End Property
Public Property Let setUSUARIO_RESULTADO2(ByVal dato As Long)
        USUARIO_RESULTADO2 = dato
End Property
Public Property Let setREALIZACION_ES_FORMADOR(ByVal dato As Integer)
        REALIZACION_ES_FORMADOR = dato
End Property
Public Property Let setOBSERVADOR_ES_FORMADOR(ByVal dato As Integer)
        OBSERVADOR_ES_FORMADOR = dato
End Property
Public Property Let setES_DUPLICADO(ByVal dato As Integer)
        ES_DUPLICADO = dato
End Property
Public Property Let setES_RECUALIFICACION(ByVal dato As Integer)
        ES_RECUALIFICACION = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMUESTRAS_OBSERVADORES
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getDETERMINACION_ENSAYO_ID() As Long
        getDETERMINACION_ENSAYO_ID = DETERMINACION_ENSAYO_ID
End Property
Public Property Get getTIPO_DETERMINACION_ENSAYO_ID() As Long
        getTIPO_DETERMINACION_ENSAYO_ID = TIPO_DETERMINACION_ENSAYO_ID
End Property
'M0961-I
Public Property Get getSELLANTE_ID() As Long
        getSELLANTE_ID = SELLANTE_ID
End Property
Public Property Get getENSAYO() As Long
        getENSAYO = ENSAYO
End Property
'M0961-F
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getES_CONTROL_EFICACIA() As Integer
        getES_CONTROL_EFICACIA = ES_CONTROL_EFICACIA
End Property
Public Property Get getUSUARIO_ID_REALIZACION() As Long
        getUSUARIO_ID_REALIZACION = USUARIO_ID_REALIZACION
End Property
Public Property Get getUSUARIO_ID_OBSERVADOR() As Long
        getUSUARIO_ID_OBSERVADOR = USUARIO_ID_OBSERVADOR
End Property
Public Property Get getUSUARIO_RESULTADO1() As Long
        getUSUARIO_RESULTADO1 = USUARIO_RESULTADO1
End Property
Public Property Get getUSUARIO_RESULTADO2() As Long
        getUSUARIO_RESULTADO2 = USUARIO_RESULTADO2
End Property
Public Property Get getREALIZACION_ES_FORMADOR() As Integer
        getREALIZACION_ES_FORMADOR = REALIZACION_ES_FORMADOR
End Property
Public Property Get getOBSERVADOR_ES_FORMADOR() As Integer
        getOBSERVADOR_ES_FORMADOR = OBSERVADOR_ES_FORMADOR
End Property
Public Property Get getES_DUPLICADO() As Integer
        getES_DUPLICADO = ES_DUPLICADO
End Property
Public Property Get getES_RECUALIFICACION() As Integer
        getES_RECUALIFICACION = ES_RECUALIFICACION
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMUESTRAS_OBSERVADORES
'***************************************************************
Public Function Carga(MUESTRA As Long, DETERMINACION As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MUESTRAS_OBSERVADORES " & _
                   " WHERE MUESTRA_ID = " & MUESTRA & " AND DETERMINACION_ENSAYO_ID=" & DETERMINACION
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                DETERMINACION_ENSAYO_ID = rs("DETERMINACION_ENSAYO_ID")
                TIPO_DETERMINACION_ENSAYO_ID = rs("TIPO_DETERMINACION_ENSAYO_ID")
                'M0961-I
                SELLANTE_ID = rs("SELLANTE_ID")
                ENSAYO = rs("ENSAYO")
                'M0961-F
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                fecha = rs("FECHA")
                ES_CONTROL_EFICACIA = rs("ES_CONTROL_EFICACIA")
                USUARIO_ID_REALIZACION = rs("USUARIO_ID_REALIZACION")
                USUARIO_ID_OBSERVADOR = rs("USUARIO_ID_OBSERVADOR")
                USUARIO_RESULTADO1 = rs("USUARIO_RESULTADO1")
                USUARIO_RESULTADO2 = rs("USUARIO_RESULTADO2")
                REALIZACION_ES_FORMADOR = rs("REALIZACION_ES_FORMADOR")
                OBSERVADOR_ES_FORMADOR = rs("OBSERVADOR_ES_FORMADOR")
                ES_DUPLICADO = rs("ES_DUPLICADO")
                ES_RECUALIFICACION = rs("ES_RECUALIFICACION")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMuestras_observadores)", vbCritical, Err.Description
End Function
'M0961-I
Public Function CargaSellante(MUESTRA As Long, SELLANTE_ID As Long, ENSAYO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * FROM MUESTRAS_OBSERVADORES " & _
                   " WHERE MUESTRA_ID = " & MUESTRA & " AND SELLANTE_ID=" & SELLANTE_ID & " AND ENSAYO=" & ENSAYO
                   
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                CargaSellante = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                DETERMINACION_ENSAYO_ID = rs("DETERMINACION_ENSAYO_ID")
                TIPO_DETERMINACION_ENSAYO_ID = rs("TIPO_DETERMINACION_ENSAYO_ID")
                SELLANTE_ID = rs("SELLANTE_ID")
                ENSAYO = rs("ENSAYO")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                fecha = rs("FECHA")
                ES_CONTROL_EFICACIA = rs("ES_CONTROL_EFICACIA")
                USUARIO_ID_REALIZACION = rs("USUARIO_ID_REALIZACION")
                USUARIO_ID_OBSERVADOR = rs("USUARIO_ID_OBSERVADOR")
                USUARIO_RESULTADO1 = rs("USUARIO_RESULTADO1")
                USUARIO_RESULTADO2 = rs("USUARIO_RESULTADO2")
                REALIZACION_ES_FORMADOR = rs("REALIZACION_ES_FORMADOR")
                OBSERVADOR_ES_FORMADOR = rs("OBSERVADOR_ES_FORMADOR")
                ES_DUPLICADO = rs("ES_DUPLICADO")
                ES_RECUALIFICACION = rs("ES_RECUALIFICACION")
                OBSERVACIONES = rs("OBSERVACIONES")
                CargaSellante = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaSellante = False
        MsgBox "Error al cargar los datos(clsMuestras_observadores)", vbCritical, Err.Description
End Function


'M0961-f

Public Function Insertar() As Long
'M0961: Añade los campos SELLANTE_ID y ENSAYO
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO MUESTRAS_OBSERVADORES " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & DETERMINACION_ENSAYO_ID & "," & TIPO_DETERMINACION_ENSAYO_ID & "," & DOCUMENTO_ID & "," & _
                        SELLANTE_ID & "," & ENSAYO & "," & "'" & fecha & "'" & "," & ES_CONTROL_EFICACIA & "," & _
                        USUARIO_ID_REALIZACION & "," & USUARIO_ID_OBSERVADOR & "," & _
                        USUARIO_RESULTADO1 & "," & USUARIO_RESULTADO2 & "," & _
                        REALIZACION_ES_FORMADOR & "," & _
                        OBSERVADOR_ES_FORMADOR & "," & ES_DUPLICADO & "," & ES_RECUALIFICACION & "," & _
                        "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMuestras_observadores)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MUESTRAS_OBSERVADORES SET " & _
                        " DETERMINACION_ENSAYO_ID = " & DETERMINACION_ENSAYO_ID & "," & _
                        " TIPO_DETERMINACION_ENSAYO_ID = " & TIPO_DETERMINACION_ENSAYO_ID & "," & _
                        " SELLANTE_ID = " & SELLANTE_ID & "," & _
                        " ENSAYO = " & ENSAYO & "," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " ES_CONTROL_EFICACIA = " & ES_CONTROL_EFICACIA & "," & _
                        " USUARIO_ID_REALIZACION = " & USUARIO_ID_REALIZACION & "," & _
                        " USUARIO_ID_OBSERVADOR = " & USUARIO_ID_OBSERVADOR & "," & _
                        " USUARIO_RESULTADO1 = " & USUARIO_RESULTADO1 & "," & _
                        " USUARIO_RESULTADO2 = " & USUARIO_RESULTADO2 & "," & _
                        " REALIZACION_ES_FORMADOR = " & REALIZACION_ES_FORMADOR & "," & _
                        " OBSERVADOR_ES_FORMADOR = " & OBSERVADOR_ES_FORMADOR & "," & _
                        " ES_DUPLICADO = " & ES_DUPLICADO & "," & _
                        " ES_RECUALIFICACION = " & ES_RECUALIFICACION & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMuestras_observadores)", vbCritical, Err.Description
End Function
Public Function Eliminar(MUESTRA As Long, DETERMINACION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MUESTRAS_OBSERVADORES " & _
                " WHERE MUESTRA_ID = " & MUESTRA & " AND DETERMINACION_ENSAYO_ID=" & DETERMINACION
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMuestras_observadores)", vbCritical, Err.Description
End Function
'M0961-I
Public Function EliminarSellante(MUESTRA As Long, SELLANTE_ID As Long, ENSAYO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MUESTRAS_OBSERVADORES " & _
                " WHERE MUESTRA_ID = " & MUESTRA & " AND SELLANTE_ID=" & SELLANTE_ID & " AND ENSAYO=" & ENSAYO
        execute_bd consulta
        EliminarSellante = True
        Exit Function
fallo:
        EliminarSellante = False
        MsgBox "Error al Eliminar (clsMuestras_observadores)", vbCritical, Err.Description
End Function
'M0961-F
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MUESTRAS_OBSERVADORES ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function
