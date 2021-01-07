VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarifas_portes_articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTARIFAS_PORTES_ARTICULOS
'***************************************************************
Private TARIFA_PORTE_ID As Long
Private ARTICULO_ID As Long
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTARIFAS_PORTES_ARTICULOS
'***************************************************************
Public Property Let setTARIFA_PORTE_ID(ByVal dato As Long)
        TARIFA_PORTE_ID = dato
End Property
Public Property Let setARTICULO_ID(ByVal dato As Long)
        ARTICULO_ID = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTARIFAS_PORTES_ARTICULOS
'***************************************************************
Public Property Get getTARIFA_PORTE_ID() As Long
        getTARIFA_PORTE_ID = TARIFA_PORTE_ID
End Property
Public Property Get getARTICULO_ID() As Long
        getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTARIFAS_PORTES_ARTICULOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS_PORTES_ARTICULOS WHERE TARIFA_PORTE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TARIFA_PORTE_ID = rs("TARIFA_PORTE_ID")
                ARTICULO_ID = rs("ARTICULO_ID")
                PRECIO = rs("PRECIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTarifas_portes_articulos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TARIFAS_PORTES_ARTICULOS " & _
                   "  VALUES (" & _
                        TARIFA_PORTE_ID & "," & ARTICULO_ID & "," & "'" & PRECIO & "'" & _
                ")"
        execute_bd consulta
        Insertar = TARIFA_PORTE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTarifas_portes_articulos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long, ARTICULO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TARIFAS_PORTES_ARTICULOS SET " & _
                        " PRECIO = '" & PRECIO & "'" & _
                " WHERE TARIFA_PORTE_ID = " & ID & _
                "   AND ARTICULO_ID = " & ARTICULO

        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTarifas_portes_articulos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TARIFAS_PORTES_ARTICULOS " & _
                "    WHERE ARTICULO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTarifas_portes_articulos)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ARTICULO_ID, B.DESCRIPCION, A.PRECIO " & _
                   "  FROM TARIFAS_PORTES_ARTICULOS A, ARTICULOS B " & _
                   " WHERE A.TARIFA_PORTE_ID = " & ID & _
                   "   AND A.ARTICULO_ID = B.ID_ARTICULO " & _
                   " ORDER BY A.ARTICULO_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_articulo(ARTICULO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS_PORTES_ARTICULOS WHERE ARTICULO_ID = " & ARTICULO & " ORDER BY TARIFA_PORTE_ID"
        Set Listado_por_articulo = datos_bd(consulta)
End Function

