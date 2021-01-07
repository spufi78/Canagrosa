VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPc_rex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPC_REX
'***************************************************************
Private PEDIDO_ID As Long
Private BOTE_EX_ID As Long
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPC_REX
'***************************************************************
Public Property Let setPEDIDO_ID(ByVal dato As Long)
        PEDIDO_ID = dato
End Property
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPC_REX
'***************************************************************
Public Property Get getPEDIDO_ID() As Long
        getPEDIDO_ID = PEDIDO_ID
End Property
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPC_REX
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM PC_REX WHERE PEDIDO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PEDIDO_ID = rs("PEDIDO_ID")
                BOTE_EX_ID = rs("BOTE_EX_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPc_rex)", vbCritical, Err.Description
End Function
Public Function Insertar(fechaPedido As Date) As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PC_REX " & _
                   "  VALUES (" & _
                        PEDIDO_ID & "," & BOTE_EX_ID & "," & ORDEN & _
                ")"
        execute_bd consulta
        Dim oB As New clsBotes_ex
        oB.observacion BOTE_EX_ID, vbNewLine & "Se inserta en el pedido de producto controlado número " & PEDIDO_ID & " el día " & Format(Date, "dd-mm-yyyy")
        oB.Terminar BOTE_EX_ID, CStr(fechaPedido)
        Set oB = Nothing
        Insertar = PEDIDO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPc_rex)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PC_REX SET " & _
                        " BOTE_EX_ID = " & BOTE_EX_ID & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE PEDIDO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPc_rex)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim oB As New clsBotes_ex
        Dim rs As ADODB.RecordSet
        Set rs = datos_bd("SELECT * FROM PC_REX WHERE PEDIDO_ID = " & ID)
        If rs.RecordCount > 0 Then
            Do
                oB.observacion rs("BOTE_EX_ID"), vbNewLine & "Se elimina del pedido de producto controlado número " & ID & " el día " & Format(Date, "dd-mm-yyyy")
                oB.AnularTerminacion rs("BOTE_EX_ID")
                rs.MoveNext
            Loop Until rs.EOF
        End If
        consulta = "DELETE FROM PC_REX WHERE PEDIDO_ID = " & ID
        execute_bd consulta
        Set oB = Nothing
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPc_rex)", vbCritical, Err.Description
End Function
Public Function Listado(PEDIDO As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT A.ID_BOTE_EX, " & _
                   "       C.NOMBRE, " & _
                   "       A.FECHA_PEDIDO, " & _
                   "       A.FECHA_CADUCIDAD, " & _
                   "       A.LOTE " & _
                   " FROM BOTES_EX A, " & _
                   "      TIPOS_BOTE_EX B, " & _
                   "      TIPOS_REACTIVO_EX C, " & _
                   "      PC_REX PCR " & _
                   " WHERE A.TIPO_BOTE_EX_ID = B.ID_TIPO_BOTE_EX " & _
                   "   AND B.TIPO_REACTIVO_EX_ID = C.ID_TIPO_REACTIVO_EX " & _
                   "   AND A.ID_BOTE_EX = PCR.BOTE_EX_ID " & _
                   "   AND PCR.PEDIDO_ID = " & PEDIDO & _
                   " ORDER BY A.ID_BOTE_EX ASC"
        Set Listado = datos_bd(consulta)
End Function
