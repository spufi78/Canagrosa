VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestras_soluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMUESTRAS_SOLUCIONES
'***************************************************************
Private MUESTRA_ID As Long
Private fecha_fabricacion As String
Private fecha_caducidad As String
Private ETIQUETA_ID As Long
Private TAMANO As Long
Private SP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMUESTRAS_SOLUCIONES
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setFECHA_FABRICACION(ByVal dato As String)
        fecha_fabricacion = dato
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
        fecha_caducidad = dato
End Property
Public Property Let setETIQUETA_ID(ByVal dato As Long)
        ETIQUETA_ID = dato
End Property
Public Property Let setTAMANO(ByVal dato As Long)
        TAMANO = dato
End Property
Public Property Let setSP(ByVal dato As String)
        SP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMUESTRAS_SOLUCIONES
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getFECHA_FABRICACION() As String
        getFECHA_FABRICACION = fecha_fabricacion
End Property
Public Property Get getFECHA_CADUCIDAD() As String
        getFECHA_CADUCIDAD = fecha_caducidad
End Property
Public Property Get getETIQUETA_ID() As Long
        getETIQUETA_ID = ETIQUETA_ID
End Property
Public Property Get getTAMANO() As Long
        getTAMANO = TAMANO
End Property
Public Property Get getSP() As String
        getSP = SP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMUESTRAS_SOLUCIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MUESTRAS_SOLUCIONES WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                fecha_fabricacion = rs("FECHA_FABRICACION")
                fecha_caducidad = rs("FECHA_CADUCIDAD")
                ETIQUETA_ID = rs("ETIQUETA_ID")
                TAMANO = rs("TAMANO")
                SP = rs("SP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMuestras_soluciones)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE INTO MUESTRAS_SOLUCIONES " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & "'" & Format(fecha_fabricacion, "yyyy-mm-dd") & "'" & "," & "'" & Format(fecha_caducidad, "yyyy-mm-dd") & "'" & "," & _
                        ETIQUETA_ID & "," & TAMANO & ",'" & SP & "'" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMuestras_soluciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MUESTRAS_SOLUCIONES " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMuestras_soluciones)", vbCritical, Err.Description
End Function
