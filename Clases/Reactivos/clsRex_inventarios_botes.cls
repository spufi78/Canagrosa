VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRex_inventarios_botes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSREX_INVENTARIOS_BOTES
'***************************************************************
Private INVENTARIO_ID As Long
Private BOTE_EX_ID As Long
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSREX_INVENTARIOS_BOTES
'***************************************************************
Public Property Let setINVENTARIO_ID(ByVal dato As Long)
        INVENTARIO_ID = dato
End Property
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSREX_INVENTARIOS_BOTES
'***************************************************************
Public Property Get getINVENTARIO_ID() As Long
        getINVENTARIO_ID = INVENTARIO_ID
End Property
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSREX_INVENTARIOS_BOTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REX_INVENTARIOS_BOTES WHERE INVENTARIO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                INVENTARIO_ID = rs("INVENTARIO_ID")
                BOTE_EX_ID = rs("BOTE_EX_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRex_inventarios_botes)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO REX_INVENTARIOS_BOTES " & _
                   "  VALUES (" & _
                        INVENTARIO_ID & "," & BOTE_EX_ID & "," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = INVENTARIO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRex_inventarios_botes)", vbCritical, Err.Description
End Function
Public Function InsertarCadena(INVENTARIO_ID As Long, CADENA As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim cad() As String
        If CADENA <> "" Then
            cad = Split(CADENA, ",")
        End If
        Dim c As String
        Dim v As String
        c = "insert into rex_inventarios_botes values "
        v = ""
        Dim i As Integer
        For i = LBound(cad) To UBound(cad)
            If v <> "" Then
                v = v & ","
            End If
            v = v & "(" & INVENTARIO_ID & "," & cad(i) & "," & i & ")"
            If i <> 0 And i Mod 50 = 0 Then
                execute_bd c & v
                v = ""
            End If
        Next
        If v <> "" Then
            execute_bd c & v
        End If
        InsertarCadena = True
        Exit Function
fallo:
        InsertarCadena = False
        MsgBox "Error al InsertarCadena (clsRex_inventarios_botes)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE REX_INVENTARIOS_BOTES SET " & _
                        " BOTE_EX_ID = " & BOTE_EX_ID & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE INVENTARIO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRex_inventarios_botes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM REX_INVENTARIOS_BOTES " & _
                "    WHERE INVENTARIO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRex_inventarios_botes)", vbCritical, Err.Description
End Function
Public Function Listado(PK As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT be.id_bote_ex, " & _
                   "       tr.nombre, " & _
                   "       be.LOTE, " & _
                   "       be.fecha_caducidad " & _
                   " FROM BOTES_EX be, " & _
                   "      TIPOS_BOTE_EX tb, " & _
                   "      TIPOS_REACTIVO_EX tr, " & _
                   "      REX_INVENTARIOS_BOTES ib " & _
                   " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex " & _
                   "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
                   "   AND be.id_bote_ex = ib.bote_ex_id " & _
                   "   AND ib.inventario_id = " & PK & _
                   " ORDER BY tr.nombre asc"
        Set Listado = datos_bd(consulta)
End Function
