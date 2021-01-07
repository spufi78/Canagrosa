VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarifas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTARIFAS
'***************************************************************
Private OBRA_ID As Long
Private ARTICULO_ID As Long
Private PRECIO_FABRICA As String
Private PRECIO_OBRA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTARIFAS
'***************************************************************
Public Property Let setOBRA_ID(ByVal dato As Long)
        OBRA_ID = dato
End Property
Public Property Let setARTICULO_ID(ByVal dato As Long)
        ARTICULO_ID = dato
End Property
Public Property Let setPRECIO_FABRICA(ByVal dato As String)
        PRECIO_FABRICA = dato
End Property
Public Property Let setPRECIO_OBRA(ByVal dato As String)
        PRECIO_OBRA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTARIFAS
'***************************************************************
Public Property Get getOBRA_ID() As Long
        getOBRA_ID = OBRA_ID
End Property
Public Property Get getARTICULO_ID() As Long
        getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getPRECIO_FABRICA() As String
        getPRECIO_FABRICA = PRECIO_FABRICA
End Property
Public Property Get getPRECIO_OBRA() As String
        getPRECIO_OBRA = PRECIO_OBRA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTARIFAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS WHERE OBRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                OBRA_ID = rs("OBRA_ID")
                ARTICULO_ID = rs("ARTICULO_ID")
                PRECIO_FABRICA = rs("PRECIO_FABRICA")
                PRECIO_OBRA = rs("PRECIO_OBRA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTarifas)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TARIFAS " & _
                   "  VALUES (" & _
                        OBRA_ID & "," & ARTICULO_ID & ",'" & PRECIO_FABRICA & "','" & _
                        PRECIO_OBRA & "'" & _
                ")"
        execute_bd consulta
        Insertar = OBRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTarifas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TARIFAS SET " & _
                        " ARTICULO_ID = " & ARTICULO_ID & "," & _
                        " PRECIO_FABRICA = '" & PRECIO_FABRICA & "'," & _
                        " PRECIO_OBRA = '" & PRECIO_OBRA & "'" & _
                " WHERE OBRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTarifas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TARIFAS " & _
                "    WHERE OBRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTarifas)", vbCritical, Err.Description
End Function
Public Function Listado(pk As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.ID_ARTICULO, B.DESCRIPCION, A.PRECIO_FABRICA, A.PRECIO_OBRA,TP.PRECIO " & _
                   "  FROM TARIFAS A " & _
                   " INNER JOIN ARTICULOS B ON A.ARTICULO_ID = B.ID_ARTICULO " & _
                   " INNER JOIN OBRAS O ON O.ID_OBRA = A.OBRA_ID " & _
                   " LEFT JOIN TARIFAS_PORTES_ARTICULOS TP ON O.TARIFA_PORTE_ID = TP.TARIFA_PORTE_ID AND A.ARTICULO_ID = TP.ARTICULO_ID " & _
                   " WHERE A.OBRA_ID = " & pk & _
                   " ORDER BY B.ID_ARTICULO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Completo() As ADODB.Recordset
        Dim consulta As String
'        consulta = "SELECT DISTINCT B.ID_ARTICULO, B.DESCRIPCION, 0, 0,TP.PRECIO " & _
'                   "  FROM ARTICULOS B,OBRAS O, TARIFAS_PORTES_ARTICULOS TP " & _
'                   " WHERE O.ID_OBRA = " & PK & _
'                   "   AND O.TARIFA_PORTE_ID = TP.TARIFA_PORTE_ID " & _
'                   "   AND B.ID_ARTICULO = TP.ARTICULO_ID " & _
'                   " ORDER BY B.ID_ARTICULO"
        consulta = "SELECT DISTINCT B.ID_ARTICULO, B.DESCRIPCION, 0, 0, 0 " & _
                   "  FROM ARTICULOS B " & _
                   " ORDER BY B.ID_ARTICULO"
        Set Listado_Completo = datos_bd(consulta)
End Function


