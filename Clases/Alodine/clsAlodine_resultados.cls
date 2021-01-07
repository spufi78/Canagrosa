VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsALODINE_RESULTADOS
'***************************************************************
Private LOTE_ID As Long
Private ALODINE_ID As Long
Private PARAMETRO_ID As Long
Private RESULTADO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsALODINE_RESULTADOS
'***************************************************************
Public Property Let setLOTE_ID(ByVal dato As Long)
        LOTE_ID = dato
End Property
Public Property Let setALODINE_ID(ByVal dato As Long)
        ALODINE_ID = dato
End Property
Public Property Let setPARAMETRO_ID(ByVal dato As Long)
        PARAMETRO_ID = dato
End Property
Public Property Let setRESULTADO(ByVal dato As String)
        RESULTADO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsALODINE_RESULTADOS
'***************************************************************
Public Property Get getLOTE_ID() As Long
        getLOTE_ID = LOTE_ID
End Property
Public Property Get getALODINE_ID() As Long
        getALODINE_ID = ALODINE_ID
End Property
Public Property Get getPARAMETRO_ID() As Long
        getPARAMETRO_ID = PARAMETRO_ID
End Property
Public Property Get getRESULTADO() As String
        getRESULTADO = RESULTADO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsALODINE_RESULTADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_RESULTADOS WHERE LOTE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                LOTE_ID = rs("LOTE_ID")
                ALODINE_ID = rs("ALODINE_ID")
                PARAMETRO_ID = rs("PARAMETRO_ID")
                RESULTADO = rs("RESULTADO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsALODINE_RESULTADOS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO ALODINE_RESULTADOS " & _
                   "  VALUES (" & _
                        LOTE_ID & "," & ALODINE_ID & "," & PARAMETRO_ID & "," & _
                        "'" & RESULTADO & "'" & _
                ")"
        execute_bd consulta
        Insertar = LOTE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsALODINE_RESULTADOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID_LOTE As Long, ID_PARAMETRO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_RESULTADOS SET " & _
                        " RESULTADO = '" & RESULTADO & "'" & _
                " WHERE LOTE_ID = " & ID_LOTE & _
                "   AND PARAMETRO_ID = " & ID_PARAMETRO
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsALODINE_RESULTADOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_RESULTADOS " & _
                "    WHERE LOTE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsALODINE_RESULTADOS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_RESULTADOS ORDER BY LOTE_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Resultados_Lote(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT AP.PARAMETRO, AP.RANGO, U.NOMBRE, AR.RESULTADO,AP.PROCEDIMIENTO, AP.ID_PARAMETRO " & _
                   "  FROM ALODINE_RESULTADOS AS AR, ALODINE_PARAMETROS AS AP, UNIDADES AS U " & _
                   " WHERE AR.PARAMETRO_ID = AP.ID_PARAMETRO " & _
                   "   AND AP.UNIDAD_ID = U.ID_UNIDAD " & _
                   "   AND AR.LOTE_ID = " & ID & _
                   " ORDER BY AP.ID_PARAMETRO"
        Set Resultados_Lote = datos_bd(consulta)
End Function
