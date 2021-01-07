VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_ensayos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_ENSAYOS
'***************************************************************
Private FICHA_ID As Long
Private TIPO_ENSAYO_ID As Long
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_ENSAYOS
'***************************************************************
Public Property Let setFICHA_ID(ByVal dato As Long)
        FICHA_ID = dato
End Property
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_ENSAYOS
'***************************************************************
Public Property Get getFICHA_ID() As Long
        getFICHA_ID = FICHA_ID
End Property
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_ENSAYOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_ENSAYOS WHERE FICHA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                FICHA_ID = rs("FICHA_ID")
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_ensayos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CE_ENSAYOS " & _
                   "  VALUES (" & _
                        FICHA_ID & "," & _
                        TIPO_ENSAYO_ID & "," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = FICHA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_ensayos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_ENSAYOS SET " & _
                        " TIPO_ENSAYO_ID = " & TIPO_ENSAYO_ID & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE FICHA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_ensayos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_ENSAYOS " & _
                "    WHERE FICHA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_ensayos)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.NOMBRE,C.NOMBRE,B.ID_TIPO_ENSAYO,B.TIPO_ANALISIS_ID,B.INCLUYE_ESPESOR,B.LOTE_PROBETAS,B.NORMA " & _
                   "  FROM CE_ENSAYOS A, CE_TIPOS_ENSAYOS B " & _
                   " LEFT JOIN PROCESOS_BASE C ON B.PROCESO_BASE_ID = C.ID_PROCESO_BASE " & _
                   " WHERE A.FICHA_ID = " & ID & _
                   "   AND A.TIPO_ENSAYO_ID = B.ID_TIPO_ENSAYO " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
Public Function lista(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM CE_ENSAYOS " & _
                   " WHERE FICHA_ID = " & ID
        Set lista = datos_bd(consulta)
End Function
Public Function ListadoSubcontratablesParametro(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.NOMBRE,C.NOMBRE,B.ID_TIPO_ENSAYO,B.TIPO_ANALISIS_ID,B.INCLUYE_ESPESOR,B.LOTE_PROBETAS,B.NORMA " & _
                   "  FROM CE_ENSAYOS A, CE_TIPOS_ENSAYOS B " & _
                   " LEFT JOIN PROCESOS_BASE C ON B.PROCESO_BASE_ID = C.ID_PROCESO_BASE " & _
                   " WHERE A.FICHA_ID = " & ID & _
                   "   AND A.TIPO_ENSAYO_ID = B.ID_TIPO_ENSAYO " & _
                   " ORDER BY A.ORDEN"
        Set ListadoSubcontratablesParametro = datos_bd(consulta)
End Function
