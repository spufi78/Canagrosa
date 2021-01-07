VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsARTICULOS_TARIFAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSARTICULOS_TARIFAS
'***************************************************************
Private ARTICULO_ID As Long
Private TARIFA_ID As Long
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSARTICULOS_TARIFAS
'***************************************************************
Public Property Let setARTICULO_ID(ByVal dato As Long)
        ARTICULO_ID = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setTARIFA_ID(ByVal dato As Long)
        TARIFA_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSARTICULOS_TARIFAS
'***************************************************************
Public Property Get getARTICULO_ID() As Long
        getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getTARIFA_ID() As Long
        getTARIFA_ID = TARIFA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSARTICULOS_TARIFAS
'***************************************************************
Public Function Carga(ARTICULO As Long, TARIFA As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS_TARIFAS " & _
                   " WHERE ARTICULO_ID = " & ARTICULO & _
                   "   AND TARIFA_ID = " & TARIFA
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PRECIO = rs("PRECIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsARTICULOS_TARIFAS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO ARTICULOS_TARIFAS " & _
                   "  VALUES (" & _
                        ARTICULO_ID & "," & TARIFA_ID & ",'" & PRECIO & "'" & _
                ")"
        conn.Execute consulta
        Insertar = ARTICULO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsARTICULOS_TARIFAS)", vbCritical, Err.Description
End Function
Public Function Modificar(ARTICULO As Long, TARIFA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ARTICULOS_TARIFAS SET " & _
                        " PRECIO = '" & PRECIO & "'" & _
                   " WHERE ARTICULO_ID = " & ARTICULO & _
                   "   AND TARIFA_ID = " & TARIFA
        conn.Execute consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsARTICULOS_TARIFAS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ARTICULOS_TARIFAS " & _
                "    WHERE ARTICULO_ID = " & ID
        conn.Execute consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsARTICULOS_TARIFAS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS_TARIFAS ORDER BY ARTICULO_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_Articulo(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS_TARIFAS " & _
                   " WHERE ARTICULO_ID = " & ID & _
                   " ORDER BY TARIFA_ID"
        Set Listado_por_Articulo = datos_bd(consulta)
End Function

