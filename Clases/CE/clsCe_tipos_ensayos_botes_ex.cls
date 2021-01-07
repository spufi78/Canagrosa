VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_tipos_ensayos_botes_ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_TIPOS_ENSAYOS_BOTES_EX
'***************************************************************
Private TIPO_ENSAYO_ID As Long
Private BOTE_EX_ID As Long
Private tipo As String
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_TIPOS_ENSAYOS_BOTES_EX
'***************************************************************
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As String)
        tipo = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_TIPOS_ENSAYOS_BOTES_EX
'***************************************************************
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getTIPO() As String
        getTIPO = tipo
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_TIPOS_ENSAYOS_BOTES_EX
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_TIPOS_ENSAYOS_BOTES_EX WHERE TIPO_ENSAYO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                BOTE_EX_ID = rs("BOTE_EX_ID")
                tipo = rs("TIPO")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCE_TIPOS_ENSAYOS_BOTES_EX)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CE_TIPOS_ENSAYOS_BOTES_EX " & _
                   "  VALUES (" & _
                        TIPO_ENSAYO_ID & "," & BOTE_EX_ID & ",'" & tipo & "'," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = TIPO_ENSAYO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCE_TIPOS_ENSAYOS_BOTES_EX)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_TIPOS_ENSAYOS_BOTES_EX " & _
                "    WHERE TIPO_ENSAYO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCE_TIPOS_ENSAYOS_BOTES_EX)", vbCritical, Err.Description
End Function
Public Function Listado(TD As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT A.BOTE_EX_ID,A.TIPO " & _
                   " FROM CE_TIPOS_ENSAYOS_BOTES_EX A " & _
                   " WHERE A.TIPO_ENSAYO_ID = " & TD & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function


