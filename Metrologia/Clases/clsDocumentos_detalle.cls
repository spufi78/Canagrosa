VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentos_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private DOCUMENTO_ID As Long
Private orden As Long
Private ARTICULO_ID As Long
Private FECHA_ALBARAN As String
Private NUMERO_ALBARAN As Long
Private SERVIDO As String
Private CANTIDAD As Long
Private descripcion As String
Private PRECIO As String
Private total As String
Private portes As String
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        orden = dato
End Property
Public Property Let setARTICULO_ID(ByVal dato As Long)
        ARTICULO_ID = dato
End Property
Public Property Let setFECHA_ALBARAN(ByVal dato As String)
        FECHA_ALBARAN = dato
End Property
Public Property Let setNUMERO_ALBARAN(ByVal dato As Long)
        NUMERO_ALBARAN = dato
End Property
Public Property Let setSERVIDO(ByVal dato As String)
        SERVIDO = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As Long)
        CANTIDAD = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        descripcion = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setTOTAL(ByVal dato As String)
        total = dato
End Property
Public Property Let setPORTES(ByVal dato As String)
        portes = dato
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = orden
End Property
Public Property Get getARTICULO_ID() As Long
        getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getFECHA_ALBARAN() As String
        getFECHA_ALBARAN = FECHA_ALBARAN
End Property
Public Property Get getNUMERO_ALBARAN() As Long
        getNUMERO_ALBARAN = NUMERO_ALBARAN
End Property
Public Property Get getSERVIDO() As String
        getSERVIDO = SERVIDO
End Property
Public Property Get getCANTIDAD() As Long
        getCANTIDAD = CANTIDAD
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = descripcion
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getTOTAL() As String
        getTOTAL = total
End Property
Public Property Get getPORTES() As String
        getPORTES = portes
End Property
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_DETALLE WHERE DOCUMENTO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                orden = rs("ORDEN")
                ARTICULO_ID = rs("ARTICULO_ID")
                FECHA_ALBARAN = rs("FECHA_ALBARAN")
                NUMERO_ALBARAN = rs("NUMERO_ALBARAN")
                SERVIDO = rs("SERVIDO")
                CANTIDAD = rs("CANTIDAD")
                descripcion = rs("DESCRIPCION")
                PRECIO = rs("PRECIO")
                total = rs("TOTAL")
                portes = rs("PORTES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDOCUMENTOS_DETALLE)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO DOCUMENTOS_DETALLE " & _
                   "  VALUES (" & _
                        DOCUMENTO_ID & "," & orden & "," & _
                        ARTICULO_ID & ",'" & FECHA_ALBARAN & "'," & NUMERO_ALBARAN & ",'" & SERVIDO & "'," & CANTIDAD & ",'" & _
                        descripcion & "'," & "'" & PRECIO & "'" & "," & _
                        "'" & total & "','" & portes & "')"
        execute_bd consulta
        Insertar = DOCUMENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDOCUMENTOS_DETALLE)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "DELETE FROM DOCUMENTOS_DETALLE " & _
                "    WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDOCUMENTOS_DETALLE)", vbCritical, Err.Description
End Function
Public Function Detalle_Documento(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT D.ARTICULO_ID,D.DESCRIPCION,D.PRECIO,D.CANTIDAD,D.TOTAL,D.PORTES,D.ORDEN,D.FECHA_ALBARAN,D.NUMERO_ALBARAN,D.SERVIDO " & _
                   "  FROM DOCUMENTOS_DETALLE D " & _
                   " WHERE D.DOCUMENTO_ID = " & ID & _
                   " ORDER BY D.ORDEN"
        Set Detalle_Documento = datos_bd(consulta)
End Function
Public Function LINEA(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM DOCUMENTOS_DETALLE " & _
                   " WHERE DOCUMENTO_ID = " & ID & _
                   " ORDER BY ORDEN"
        Set LINEA = datos_bd(consulta)
End Function

Public Function Orden_Articulo(Documento As Long, ARTICULO As Long) As Integer
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT ORDEN FROM DOCUMENTOS_DETALLE " & _
                   " WHERE DOCUMENTO_ID = " & Documento & _
                   "   AND ARTICULO_ID = " & ARTICULO
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Orden_Articulo = rs(0)
        Else
            Orden_Articulo = 0
        End If
        Exit Function
fallo:
        Orden_Articulo = 0
End Function

Public Function reordenar_Detalle(Documento As Long) As Boolean
    On Error GoTo ErrorEjecucion
    
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim consulta As String
    
    reordenar_Detalle = False
    
    Set rs = LINEA(Documento)
    
           
    For i = 0 To rs.RecordCount - 1
        consulta = "UPDATE DOCUMENTOS_DETALLE SET " & _
                    "ORDEN = " & i & _
                    " WHERE DOCUMENTO_ID = " & Documento & _
                    " AND ORDEN = " & rs("ORDEN")
    
    bd.execute_bd consulta
    rs.MoveNext
    Next
    
    Set rs = Nothing
            
    reordenar_Detalle = True
    Exit Function
ErrorEjecucion:
End Function
