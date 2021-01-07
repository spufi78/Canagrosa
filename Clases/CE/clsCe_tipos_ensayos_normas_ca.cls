VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_tipos_ensayos_normas_ca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_TIPOS_ENSAYOS_NORMAS_CA
'***************************************************************
Private TIPO_ENSAYO_ID As Long
Private NORMA_ID As Long
Private METODO As String
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_TIPOS_ENSAYOS_NORMAS_CA
'***************************************************************
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
End Property
Public Property Let setMETODO(ByVal dato As String)
        METODO = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_TIPOS_ENSAYOS_NORMAS_CA
'***************************************************************
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
End Property
Public Property Get getMETODO() As String
        getMETODO = METODO
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_TIPOS_ENSAYOS_NORMAS_CA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_TIPOS_ENSAYOS_NORMAS_CA WHERE TIPO_ENSAYO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                NORMA_ID = rs("NORMA_ID")
                METODO = rs("METODO")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCE_TIPOS_ENSAYOS_NORMAS_CA)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CE_TIPOS_ENSAYOS_NORMAS_CA " & _
                   "  VALUES (" & _
                        TIPO_ENSAYO_ID & "," & NORMA_ID & ",'" & METODO & "'," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = TIPO_ENSAYO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCE_TIPOS_ENSAYOS_NORMAS_CA)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_TIPOS_ENSAYOS_NORMAS_CA " & _
                "    WHERE TIPO_ENSAYO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCE_TIPOS_ENSAYOS_NORMAS_CA)", vbCritical, Err.Description
End Function
Public Function Listado(TD As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.NORMA_ID,B.NOMBRE,B.CODIGO,B.EDICION,A.METODO,IF(B.ENAC=0,'*','') " & _
                   " FROM CE_TIPOS_ENSAYOS_NORMAS_CA A, CA_NORMAS B " & _
                   " WHERE A.TIPO_ENSAYO_ID = " & TD & _
                   "   AND A.NORMA_ID = B.ID_NORMA " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function


