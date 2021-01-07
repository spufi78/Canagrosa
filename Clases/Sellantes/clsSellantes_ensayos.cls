VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSellantes_ensayos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSELLANTES_ENSAYOS
'***************************************************************
Private SELLANTE_ID As Long
Private ORDEN As Long
Private ACTIVO As Integer
Private ENSAYO As String
Private ENSAYO_INGLES As String
Private RANGO_INFERIOR As String
Private RANGO_SUPERIOR As String
Private REFERENCIA As String
Private criterio As String
Private NORMA_CRITERIO As String
Private UNIDAD_ID As Long
Private TIPO_DETERMINACION_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSELLANTES_ENSAYOS
'***************************************************************
Public Property Let setSELLANTE_ID(ByVal dato As Long)
        SELLANTE_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setACTIVO(ByVal dato As Integer)
        ACTIVO = dato
End Property
Public Property Let setENSAYO(ByVal dato As String)
        ENSAYO = dato
End Property
Public Property Let setENSAYO_INGLES(ByVal dato As String)
        ENSAYO_INGLES = dato
End Property
Public Property Let setRANGO_INFERIOR(ByVal dato As String)
        RANGO_INFERIOR = dato
End Property
Public Property Let setRANGO_SUPERIOR(ByVal dato As String)
        RANGO_SUPERIOR = dato
End Property
Public Property Let setREFERENCIA(ByVal dato As String)
        REFERENCIA = dato
End Property
Public Property Let setCRITERIO(ByVal dato As String)
        criterio = dato
End Property
Public Property Let setNORMA_CRITERIO(ByVal dato As String)
        NORMA_CRITERIO = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Long)
        UNIDAD_ID = dato
End Property
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSELLANTES_ENSAYOS
'***************************************************************
Public Property Get getSELLANTE_ID() As Long
        getSELLANTE_ID = SELLANTE_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getACTIVO() As Integer
        getACTIVO = ACTIVO
End Property
Public Property Get getENSAYO() As String
        getENSAYO = ENSAYO
End Property
Public Property Get getENSAYO_INGLES() As String
        getENSAYO_INGLES = ENSAYO_INGLES
End Property
Public Property Get getRANGO_INFERIOR() As String
        getRANGO_INFERIOR = RANGO_INFERIOR
End Property
Public Property Get getRANGO_SUPERIOR() As String
        getRANGO_SUPERIOR = RANGO_SUPERIOR
End Property
Public Property Get getREFERENCIA() As String
        getREFERENCIA = REFERENCIA
End Property
Public Property Get getCRITERIO() As String
        getCRITERIO = criterio
End Property
Public Property Get getNORMA_CRITERIO() As String
        getNORMA_CRITERIO = NORMA_CRITERIO
End Property
Public Property Get getUNIDAD_ID() As Long
        getUNIDAD_ID = UNIDAD_ID
End Property
Public Property Get getTIPO_DETERMINACION_ID() As Long
        getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSELLANTES_ENSAYOS
'***************************************************************
Public Function Carga(SELLANTE As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM SELLANTES_ENSAYOS " & _
                   " WHERE SELLANTE_ID = " & SELLANTE & _
                   "   AND ORDEN = " & ORDEN
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                SELLANTE_ID = rs("SELLANTE_ID")
                ORDEN = rs("ORDEN")
                ACTIVO = rs("ACTIVO")
                ENSAYO = rs("ENSAYO")
                ENSAYO_INGLES = rs("ENSAYO_INGLES")
                RANGO_INFERIOR = rs("RANGO_INFERIOR")
                RANGO_SUPERIOR = rs("RANGO_SUPERIOR")
                REFERENCIA = rs("REFERENCIA")
                criterio = rs("CRITERIO")
                NORMA_CRITERIO = rs("NORMA_CRITERIO")
                UNIDAD_ID = rs("UNIDAD_ID")
                TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSellantes_ensayos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "INSERT INTO SELLANTES_ENSAYOS " & _
                   "  VALUES (" & _
                        SELLANTE_ID & "," & ORDEN & "," & ACTIVO & "," & "'" & ENSAYO & "'" & "," & _
                        "'" & ENSAYO_INGLES & "'" & "," & "'" & RANGO_INFERIOR & "'" & "," & "'" & RANGO_SUPERIOR & "','" & REFERENCIA & "','" & _
                        criterio & "','" & NORMA_CRITERIO & "'," & _
                        UNIDAD_ID & "," & TIPO_DETERMINACION_ID & _
                ")"
        execute_bd CONSULTA
        Insertar = SELLANTE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSellantes_ensayos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE SELLANTES_ENSAYOS SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " ACTIVO = " & ACTIVO & "," & _
                        " ENSAYO = '" & ENSAYO & "'," & _
                        " ENSAYO_INGLES = '" & ENSAYO_INGLES & "'," & _
                        " RANGO_INFERIOR = '" & RANGO_INFERIOR & "'," & _
                        " RANGO_SUPERIOR = '" & RANGO_SUPERIOR & "'," & _
                        " REFERENCIA = '" & REFERENCIA & "'," & _
                        " CRITERIO = '" & criterio & "'," & _
                        " NORMA_CRITERIO = '" & NORMA_CRITERIO & "'," & _
                        " UNIDAD_ID = " & UNIDAD_ID & "," & _
                        " TIPO_DETERMINACION = " & TIPO_DETERMINACION_ID & "" & _
                " WHERE SELLANTE_ID = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSellantes_ensayos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM SELLANTES_ENSAYOS " & _
                "    WHERE SELLANTE_ID = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSellantes_ensayos)", vbCritical, Err.Description
End Function
Public Function Listado(SELLANTE As Long) As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT A.ORDEN,A.ENSAYO,A.ENSAYO_INGLES,A.RANGO_INFERIOR,A.RANGO_SUPERIOR,B.NOMBRE,A.UNIDAD_ID,A.TIPO_DETERMINACION_ID,A.REFERENCIA,A.CRITERIO,A.NORMA_CRITERIO,A.ACTIVO " & _
                   "  FROM SELLANTES_ENSAYOS A, UNIDADES B " & _
                   " WHERE A.UNIDAD_ID = B.ID_UNIDAD " & _
                   "   AND A.SELLANTE_ID = " & SELLANTE & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = "SELLANTE_ID"
        .setFK_VALOR = FK
        .setTABLA = "SELLANTES_ENSAYOS"
        .setDESCRIPCION = "Ensayos"
        .setPK = "ORDEN"
        .setCAMPO = "ENSAYO"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
