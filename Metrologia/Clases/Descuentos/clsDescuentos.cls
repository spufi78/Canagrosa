VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsDescuentos
'***************************************************************
Private ID_DESCUENTO As Long
Private fecha As String
Private BANCO_ID As Long
Private CONTABILIZADO As Integer
Private CONTABILIZADO_BM As String
Private USUARIO_ID As Long
Private TIMESTAMP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsDescuentos
'***************************************************************
Public Property Let setID_DESCUENTO(ByVal dato As Long)
        ID_DESCUENTO = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setBANCO_ID(ByVal dato As Long)
        BANCO_ID = dato
End Property
Public Property Let setCONTABILIZADO(ByVal dato As Long)
        CONTABILIZADO = dato
End Property
Public Property Let setCONTABILIZADO_BM(ByVal dato As Long)
        CONTABILIZADO_BM = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTIMESTAMP(ByVal dato As String)
        TIMESTAMP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsDescuentos
'***************************************************************
Public Property Get getID_DESCUENTO() As Long
        getID_DESCUENTO = ID_DESCUENTO
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getBANCO_ID() As Long
        getBANCO_ID = BANCO_ID
End Property
Public Property Get getCONTABILIZADO() As Integer
        getCONTABILIZADO = CONTABILIZADO
End Property
Public Property Get getCONTABILIZADO_BM() As String
        getCONTABILIZADO_BM = CONTABILIZADO_BM
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTIMESTAMP() As String
        getTIMESTAMP = TIMESTAMP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsDescuentos
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DESCUENTOS WHERE ID_DESCUENTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_DESCUENTO = rs("ID_DESCUENTO")
                fecha = rs("FECHA")
                BANCO_ID = rs("BANCO_ID")
                CONTABILIZADO = rs("CONTABILIZADO")
                If Not IsNull(rs("CONTABILIZADO_BM")) Then
                    CONTABILIZADO_BM = rs("CONTABILIZADO_BM")
                End If
                USUARIO_ID = rs("USUARIO_ID")
                TIMESTAMP = rs("TIMESTAMP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDescuentos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_DESCUENTO) FROM DESCUENTOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_DESCUENTO = 1
        Else
                ID_DESCUENTO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_DESCUENTO = 0 Then
            Me.CrearID
        End If
        consulta = "INSERT INTO DESCUENTOS " & _
                   "  VALUES (" & _
                        ID_DESCUENTO & "," & "'" & fecha & "'" & "," & BANCO_ID & "," & CONTABILIZADO & ",'0000-00-00 00:00:00'," & _
                        USUARIO_ID & "," & " current_timestamp " & _
                ")"
        execute_bd consulta
        Insertar = ID_DESCUENTO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDescuentos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DESCUENTOS SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " BANCO_ID = " & BANCO_ID & "," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TIMESTAMP = current_timestamp " & _
                " WHERE ID_DESCUENTO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDescuentos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "DELETE FROM DESCUENTOS WHERE ID_DESCUENTO = " & ID
        execute_bd "DELETE FROM DESCUENTOS_DOCUMENTOS WHERE REMESA_ID = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDescuentos)", vbCritical, Err.Description
End Function
Public Function Contabilizar(ID As Long) As Boolean
        On Error GoTo fallo
        execute_bd "UPDATE DESCUENTOS SET CONTABILIZADO = 1 WHERE ID_DESCUENTO = " & ID
        Contabilizar = True
        Exit Function
fallo:
        Contabilizar = False
        MsgBox "Error al Contabilizar (clsDescuentos)", vbCritical, Err.Description
End Function
Public Function Contabilizar_BM(ID As Long) As Boolean
        On Error GoTo fallo
        execute_bd "UPDATE DESCUENTOS SET CONTABILIZADO_BM = CURRENT_TIMESTAMP WHERE ID_DESCUENTO = " & ID
        Contabilizar_BM = True
        Exit Function
fallo:
        Contabilizar_BM = False
        MsgBox "Error al Contabilizar (clsDescuentos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DESCUENTOS ORDER BY ID_DESCUENTO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_DESCUENTO,FECHA FROM DESCUENTOS ORDER BY FECHA"
        Set Listado_Combo = datos_bd(consulta)
End Function
