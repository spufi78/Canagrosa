VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanos_frecuencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSBANOS_FRECUENCIAS
'***************************************************************
Private BANO_ID As Long
Private TIPO_FRECUENCIA_ID As Long
Private CODIGO_ID As Long
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSBANOS_FRECUENCIAS
'***************************************************************
Public Property Let setBANO_ID(ByVal dato As Long)
        BANO_ID = dato
End Property
Public Property Let setTIPO_FRECUENCIA_ID(ByVal dato As Long)
        TIPO_FRECUENCIA_ID = dato
End Property
Public Property Let setCODIGO_ID(ByVal dato As Long)
        CODIGO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSBANOS_FRECUENCIAS
'***************************************************************
Public Property Get getBANO_ID() As Long
        getBANO_ID = BANO_ID
End Property
Public Property Get getTIPO_FRECUENCIA_ID() As Long
        getTIPO_FRECUENCIA_ID = TIPO_FRECUENCIA_ID
End Property
Public Property Get getCODIGO_ID() As Long
        getCODIGO_ID = CODIGO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSBANOS_FRECUENCIAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM BANOS_FRECUENCIAS WHERE BANO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                BANO_ID = rs("BANO_ID")
                TIPO_FRECUENCIA_ID = rs("TIPO_FRECUENCIA_ID")
                CODIGO_ID = rs("CODIGO_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsBanos_frecuencias)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO BANOS_FRECUENCIAS " & _
                   "  VALUES (" & _
                        BANO_ID & "," & TIPO_FRECUENCIA_ID & "," & IIf(CODIGO_ID = "0", "null", CODIGO_ID) & "," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = BANO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsBanos_frecuencias)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE BANOS_FRECUENCIAS SET " & _
                        " TIPO_FRECUENCIA_ID = " & TIPO_FRECUENCIA_ID & "," & _
                        " CODIGO_ID = " & CODIGO_ID & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE BANO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsBanos_frecuencias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM BANOS_FRECUENCIAS " & _
                "    WHERE BANO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsBanos_frecuencias)", vbCritical, Err.Description
End Function
Public Function Listado(BANO_ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.*,CONCAT(c.codigo,' - ',C.DESCRIPCION) AS CODIGO, B.NOMBRE as TIPO_FRECUENCIA " & _
                   "  FROM banos_frecuencias A " & _
                   "  LEFT JOIN tipos_frecuencia B on a.TIPO_FRECUENCIA_ID = b.ID_TIPO_FRECUENCIA " & _
                   "  LEFT JOIN tarifas_codigos C on a.CODIGO_ID = C.ID_CODIGO " & _
                   " WHERE BANO_ID = " & BANO_ID & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function


