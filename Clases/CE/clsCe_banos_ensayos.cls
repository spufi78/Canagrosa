VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_banos_ensayos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_BANOS_ENSAYOS
'***************************************************************
Private BANO_ID As Long
Private TIPO_ENSAYO_ID As Long
Private DESIGNACION As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_BANOS_ENSAYOS
'***************************************************************
Public Property Let setBANO_ID(ByVal dato As Long)
        BANO_ID = dato
End Property
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setDESIGNACION(ByVal dato As String)
        DESIGNACION = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_BANOS_ENSAYOS
'***************************************************************
Public Property Get getBANO_ID() As Long
        getBANO_ID = BANO_ID
End Property
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getDESIGNACION() As String
        getDESIGNACION = DESIGNACION
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_BANOS_ENSAYOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_BANOS_ENSAYOS WHERE BANO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                BANO_ID = rs("BANO_ID")
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                DESIGNACION = rs("DESIGNACION")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_banos_ensayos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CE_BANOS_ENSAYOS " & _
                   "  VALUES (" & _
                        BANO_ID & "," & TIPO_ENSAYO_ID & "," & "'" & DESIGNACION & "'" & "," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = BANO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_banos_ensayos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_BANOS_ENSAYOS SET " & _
                        " TIPO_ENSAYO_ID = " & TIPO_ENSAYO_ID & "," & _
                        " DESIGNACION = '" & DESIGNACION & "'," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE BANO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_banos_ensayos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_BANOS_ENSAYOS " & _
                "    WHERE BANO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_banos_ensayos)", vbCritical, Err.Description
End Function
Public Function Listado(BANO As Long) As ADODB.Recordset
        Dim consulta As String
'M1356-I: Añadimos el campo DIAS_TRABAJO del Tipo de Análisis
'        consulta = "SELECT A.NOMBRE,A.NORMA,B.DESIGNACION,A.ID_TIPO_ENSAYO,A.INCLUYE_ESPESOR,A.LOTE_PROBETAS,A.ENAC " & _
'                   " FROM CE_BANOS_ENSAYOS B, CE_TIPOS_ENSAYOS A " & _
'                   " WHERE B.BANO_ID = " & BANO & _
'                   "   AND B.TIPO_ENSAYO_ID = A.ID_TIPO_ENSAYO " & _
'                   " ORDER BY B.ORDEN"
        consulta = "SELECT A.NOMBRE,A.NORMA,B.DESIGNACION,A.ID_TIPO_ENSAYO,A.INCLUYE_ESPESOR,A.LOTE_PROBETAS,TM.REQUIERE_PRODUCTO, TA.DIAS_TRABAJO " & _
                   " FROM CE_BANOS_ENSAYOS B, CE_TIPOS_ENSAYOS A, TIPOS_ANALISIS TA, TIPOS_MUESTRA TM " & _
                   " WHERE B.BANO_ID = " & BANO & _
                   "   AND B.TIPO_ENSAYO_ID = A.ID_TIPO_ENSAYO " & _
                   "   AND A.TIPO_ANALISIS_ID = TA.ID_TIPO_ANALISIS " & _
                   "   AND TA.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
                   " ORDER BY B.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_completo(BANO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   " FROM CE_BANOS_ENSAYOS " & _
                   " WHERE BANO_ID = " & BANO & _
                   " ORDER BY ORDEN"
        Set Listado_completo = datos_bd(consulta)
End Function


