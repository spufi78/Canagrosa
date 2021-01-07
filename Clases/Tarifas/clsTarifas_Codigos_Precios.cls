VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarifas_Codigos_Precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTARIFAS_CODIGOS_PRECIOS
'***************************************************************
Private CODIGO_ID As Long
Private TARIFA_ID As Long
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTARIFAS_CODIGOS_PRECIOS
'***************************************************************
Public Property Let setCODIGO_ID(ByVal dato As Long)
        CODIGO_ID = dato
End Property
Public Property Let setTARIFA_ID(ByVal dato As Long)
        TARIFA_ID = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTARIFAS_CODIGOS_PRECIOS
'***************************************************************
Public Property Get getCODIGO_ID() As Long
        getCODIGO_ID = CODIGO_ID
End Property
Public Property Get getTARIFA_ID() As Long
        getTARIFA_ID = TARIFA_ID
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TARIFAS_CODIGOS_PRECIOS " & _
                   "  VALUES (" & _
                        CODIGO_ID & "," & _
                        TARIFA_ID & ",'" & _
                        PRECIO & "'" & _
                ")"
        execute_bd consulta
        Insertar = TARIFA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTARIFAS_CODIGOS_PRECIOS)", vbCritical, Err.Description
End Function
Public Function Modificar(CODIGO As Long, TARIFA As Long, PRECIO As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TARIFAS_CODIGOS_PRECIOS " & _
                   "   SET PRECIO = '" & PRECIO & "' " & _
                   " WHERE CODIGO_ID = " & CODIGO & _
                   "   AND TARIFA_ID = " & TARIFA
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTARIFAS_CODIGOS_PRECIOS)", vbCritical, Err.Description
End Function
Public Function Listado(CODIGO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT T.NOMBRE,TCP.PRECIO,T.ID_TARIFA " & _
                "  FROM TARIFAS T " & _
                "  LEFT JOIN TARIFAS_CODIGOS_PRECIOS TCP ON T.ID_TARIFA = TCP.TARIFA_ID " & _
                "   AND TCP.CODIGO_ID = " & CODIGO & _
                "   AND T.EN_VIGOR = 1 " & _
                " ORDER BY 1"
        Set Listado = datos_bd(consulta)
End Function
