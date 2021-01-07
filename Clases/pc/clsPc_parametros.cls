VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPc_parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsPc_parametros
'***************************************************************
Private PEDIDO_ID As Long
Private ORDEN As Long
Private PARAMETRO As String
Private REQUISITO As String
Private VALOR As String
Private unidades As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsPc_parametros
'***************************************************************
Public Property Let setPEDIDO_ID(ByVal dato As Long)
        PEDIDO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setPARAMETRO(ByVal dato As String)
        PARAMETRO = dato
End Property
Public Property Let setREQUISITO(ByVal dato As String)
        REQUISITO = dato
End Property
Public Property Let setVALOR(ByVal dato As String)
        VALOR = dato
End Property
Public Property Let setUNIDADES(ByVal dato As String)
        unidades = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsPc_parametros
'***************************************************************
Public Property Get getPEDIDO_ID() As Long
        getPEDIDO_ID = PEDIDO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getPARAMETRO() As String
        getPARAMETRO = PARAMETRO
End Property
Public Property Get getREQUISITO() As String
        getREQUISITO = REQUISITO
End Property
Public Property Get getVALOR() As String
        getVALOR = VALOR
End Property
Public Property Get getUNIDADES() As String
        getUNIDADES = unidades
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsPc_parametros
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PC_PARAMETROS WHERE PEDIDO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PEDIDO_ID = rs("PEDIDO_ID")
                ORDEN = rs("ORDEN")
                PARAMETRO = rs("PARAMETRO")
                REQUISITO = rs("REQUISITO")
                VALOR = rs("VALOR")
                unidades = rs("UNIDADES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPc_parametros)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PC_PARAMETROS " & _
                   "  VALUES (" & _
                        PEDIDO_ID & "," & ORDEN & "," & "'" & PARAMETRO & "','" & REQUISITO & "'," & _
                        "'" & VALOR & "'" & "," & "'" & unidades & "'" & _
                ")"
        execute_bd consulta
        Insertar = PEDIDO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPc_parametros)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PC_PARAMETROS SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " PARAMETRO = '" & PARAMETRO & "'," & _
                        " REQUISITO = '" & REQUISITO & "'," & _
                        " VALOR = '" & VALOR & "'," & _
                        " UNIDADES = '" & unidades & "'" & _
                " WHERE PEDIDO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPc_parametros)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PC_PARAMETROS " & _
                "    WHERE PEDIDO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPc_parametros)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT PARAMETRO,REQUISITO,VALOR,UNIDADES " & _
                   "  FROM PC_PARAMETROS " & _
                   " WHERE PEDIDO_ID = " & ID & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function
