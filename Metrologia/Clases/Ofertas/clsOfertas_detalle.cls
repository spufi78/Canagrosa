VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOfertas_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOFERTAS_DETALLE
'***************************************************************
Private OFERTA_ID As Long
Private orden As Long
Private MATERIAL_ID As Long
Private PRECIO_FABRICA As String
Private PRECIO_OBRA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOFERTAS_DETALLE
'***************************************************************
Public Property Let setOFERTA_ID(ByVal dato As Long)
        OFERTA_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        orden = dato
End Property
Public Property Let setMATERIAL_ID(ByVal dato As Long)
        MATERIAL_ID = dato
End Property
Public Property Let setPRECIO_FABRICA(ByVal dato As String)
        PRECIO_FABRICA = dato
End Property
Public Property Let setPRECIO_OBRA(ByVal dato As String)
        PRECIO_OBRA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOFERTAS_DETALLE
'***************************************************************
Public Property Get getOFERTA_ID() As Long
        getOFERTA_ID = OFERTA_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = orden
End Property
Public Property Get getMATERIAL_ID() As Long
        getMATERIAL_ID = MATERIAL_ID
End Property
Public Property Get getPRECIO_FABRICA() As String
        getPRECIO_FABRICA = PRECIO_FABRICA
End Property
Public Property Get getPRECIO_OBRA() As String
        getPRECIO_OBRA = PRECIO_OBRA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOFERTAS_DETALLE
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OFERTAS_DETALLE WHERE OFERTA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                OFERTA_ID = rs("OFERTA_ID")
                orden = rs("ORDEN")
                MATERIAL_ID = rs("MATERIAL_ID")
                PRECIO_FABRICA = rs("PRECIO_FABRICA")
                PRECIO_OBRA = rs("PRECIO_OBRA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsOfertas_detalle)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO OFERTAS_DETALLE " & _
                   "  VALUES (" & _
                        OFERTA_ID & "," & orden & "," & MATERIAL_ID & ",'" & _
                        PRECIO_FABRICA & "','" & PRECIO_OBRA & "'" & _
                ")"
        execute_bd consulta
        Insertar = OFERTA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsOfertas_detalle)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OFERTAS_DETALLE SET " & _
                        " ORDEN = " & orden & "," & _
                        " MATERIAL_ID = " & MATERIAL_ID & "," & _
                        " PRECIO_FABRICA = '" & PRECIO_FABRICA & "'," & _
                        " PRECIO_OBRA = '" & PRECIO_OBRA & "'" & _
                " WHERE OFERTA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsOfertas_detalle)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OFERTAS_DETALLE " & _
                "    WHERE OFERTA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsOfertas_detalle)", vbCritical, Err.Description
End Function
Public Function Listado(pk As Long) As ADODB.Recordset
        Dim consulta As String
'        consulta = "SELECT A.MATERIAL_ID, B.DESCRIPCION, A.PRECIO_FABRICA, A.PRECIO_OBRA " & _
'                   "  FROM OFERTAS_DETALLE A, DECODIFICADORA B " & _
'                   " WHERE A.OFERTA_ID = " & pk & _
'                   "   AND A.MATERIAL_ID = B.VALOR " & _
'                   "   AND B.CODIGO = " & DECODIFICADORA.D_OFERTAS_MATERIALES & _
'                   " ORDER BY B.PARAMETROS, A.ORDEN"
        consulta = "SELECT A.MATERIAL_ID, B.DESCRIPCION, A.PRECIO_FABRICA, A.PRECIO_OBRA " & _
                   "  FROM OFERTAS_DETALLE A LEFT JOIN DECODIFICADORA B ON A.MATERIAL_ID = B.VALOR " & _
                   " WHERE A.OFERTA_ID = " & pk & _
                   "   AND B.CODIGO = " & DECODIFICADORA.D_OFERTAS_MATERIALES & _
                   " ORDER BY B.PARAMETROS, A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function

