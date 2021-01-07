VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_datos_analisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_DATOS_ANALISIS
'***************************************************************
Private TIPO_DATO_ID As Long
Private TIPO_ANALISIS_ID As Long
Private BANO_ID As Long
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_DATOS_ANALISIS
'***************************************************************
Public Property Let setTIPO_DATO_ID(ByVal dato As Long)
        TIPO_DATO_ID = dato
End Property
Public Property Let setTIPO_ANALISIS_ID(ByVal dato As Long)
        TIPO_ANALISIS_ID = dato
End Property
Public Property Let setBANO_ID(ByVal dato As Long)
        BANO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_DATOS_ANALISIS
'***************************************************************
Public Property Get getTIPO_DATO_ID() As Long
        getTIPO_DATO_ID = TIPO_DATO_ID
End Property
Public Property Get getTIPO_ANALISIS_ID() As Long
        getTIPO_ANALISIS_ID = TIPO_ANALISIS_ID
End Property
Public Property Get getBANO_ID() As Long
        getBANO_ID = BANO_ID
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTIPOS_DATOS_ANALISIS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_DATOS_ANALISIS WHERE TIPO_DATO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TIPO_DATO_ID = rs("TIPO_DATO_ID")
                TIPO_ANALISIS_ID = rs("TIPO_ANALISIS_ID")
                BANO_ID = rs("BANO_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTipos_datos_analisis)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TIPOS_DATOS_ANALISIS " & _
                   "  VALUES (" & _
                        TIPO_DATO_ID & "," & TIPO_ANALISIS_ID & "," & BANO_ID & "," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = TIPO_DATO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTipos_datos_analisis)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TIPOS_DATOS_ANALISIS SET " & _
                        " TIPO_ANALISIS_ID = " & TIPO_ANALISIS_ID & "," & _
                        " BANO_ID = " & BANO_ID & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE TIPO_DATO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTipos_datos_analisis)", vbCritical, Err.Description
End Function
Public Function Eliminar(ANALISIS As Long, BANO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_DATOS_ANALISIS " & _
                   " WHERE TIPO_ANALISIS_ID = " & ANALISIS & _
                   "   AND BANO_ID = " & BANO
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTipos_datos_analisis)", vbCritical, Err.Description
End Function
Public Function Listado_por_tipo_analisis(TIPO_ANALISIS As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.TIPO_DATO_ID,B.NOMBRE,B.OBLIGATORIO,C.NOMBRE,C.ID_UNIDAD,A.ORDEN " & _
                   "  FROM TIPOS_DATOS_ANALISIS A,TIPOS_DATO B,UNIDADES C " & _
                   " WHERE A.TIPO_ANALISIS_ID = " & TIPO_ANALISIS & _
                   "   AND A.TIPO_DATO_ID = B.ID_TIPO_DATO " & _
                   "   AND B.UNIDAD_ID = C.ID_UNIDAD " & _
                   " ORDER BY A.ORDEN"
        Set Listado_por_tipo_analisis = datos_bd(consulta)
End Function
Public Function Listado_por_bano(BANO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.TIPO_DATO_ID,B.NOMBRE,B.OBLIGATORIO,C.NOMBRE,C.ID_UNIDAD,A.ORDEN " & _
                   "  FROM TIPOS_DATOS_ANALISIS A,TIPOS_DATO B,UNIDADES C " & _
                   " WHERE A.BANO_ID = " & BANO & _
                   "   AND A.TIPO_DATO_ID = B.ID_TIPO_DATO " & _
                   "   AND B.UNIDAD_ID = C.ID_UNIDAD " & _
                   " ORDER BY A.ORDEN"
        Set Listado_por_bano = datos_bd(consulta)
End Function


