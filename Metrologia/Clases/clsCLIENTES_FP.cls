VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCLIENTES_FP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCLIENTES_FP
'***************************************************************
Private CLIENTE_ID As Long
Private FORMA_PAGO_ID As Long
Private orden As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCLIENTES_FP
'***************************************************************
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setFORMA_PAGO_ID(ByVal dato As Long)
        FORMA_PAGO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        orden = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCLIENTES_FP
'***************************************************************
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getFORMA_PAGO_ID() As Long
        getFORMA_PAGO_ID = FORMA_PAGO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = orden
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCLIENTES_FP
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CLIENTES_FP WHERE CLIENTE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CLIENTE_ID = rs("CLIENTE_ID")
                FORMA_PAGO_ID = rs("FORMA_PAGO_ID")
                orden = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCLIENTES_FP)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CLIENTES_FP " & _
                   "  VALUES (" & _
                        CLIENTE_ID & "," & FORMA_PAGO_ID & "," & orden & _
                ")"
        execute_bd consulta
        Insertar = CLIENTE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCLIENTES_FP)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CLIENTES_FP SET " & _
                        " FORMA_PAGO_ID = " & FORMA_PAGO_ID & "," & _
                        " ORDEN = " & orden & "" & _
                " WHERE CLIENTE_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCLIENTES_FP)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CLIENTES_FP " & _
                "    WHERE CLIENTE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCLIENTES_FP)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT FP.ID_FORMA_PAGO AS C1, FP.NOMBRE AS C2 " & _
                   " FROM CLIENTES_FP CFP,FORMA_PAGO FP " & _
                   " WHERE CFP.CLIENTE_ID = " & ID & _
                   "   AND CFP.FORMA_PAGO_ID = FP.ID_FORMA_PAGO " & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function
