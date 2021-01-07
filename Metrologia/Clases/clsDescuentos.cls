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
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsDESCUENTOS
'***************************************************************
Private ID_DESCUENTO As Integer
Private NOMBRE As String
Private DTO1 As Single
Private DTO2 As Single
Private COMISION As Single
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsDESCUENTOS
'***************************************************************
Public Property Let setID_DESCUENTO(ByVal dato As Integer)
        ID_DESCUENTO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setDTO1(ByVal dato As Single)
        DTO1 = dato
End Property
Public Property Let setDTO2(ByVal dato As Single)
        DTO2 = dato
End Property
Public Property Let setCOMISION(ByVal dato As Single)
        COMISION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsDESCUENTOS
'***************************************************************
Public Property Get getID_DESCUENTO() As Integer
        getID_DESCUENTO = ID_DESCUENTO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getDTO1() As Single
        getDTO1 = DTO1
End Property
Public Property Get getDTO2() As Single
        getDTO2 = DTO2
End Property
Public Property Get getCOMISION() As Single
        getCOMISION = COMISION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsDESCUENTOS
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DESCUENTOS WHERE ID_DESCUENTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_DESCUENTO = rs("ID_DESCUENTO")
                NOMBRE = rs("NOMBRE")
                DTO1 = rs("DTO1")
                DTO2 = rs("DTO2")
                COMISION = rs("COMISION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDESCUENTOS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO DESCUENTOS " & _
                   "  VALUES (" & _
                        ID_DESCUENTO & "," & "'" & NOMBRE & "'," & _
                        DTO1 & "," & DTO2 & "," & COMISION & _
                ")"
        execute_bd consulta
        Insertar = ID_DESCUENTO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDESCUENTOS)", vbCritical, Err.Description
End Function
Public Function modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DESCUENTOS SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " DTO1 = " & DTO1 & "," & _
                        " DTO2 = " & DTO2 & "," & _
                        " COMISION = " & COMISION & _
                " WHERE ID_DESCUENTO = " & ID
        execute_bd consulta
        modificar = True
        Exit Function
fallo:
        modificar = False
        MsgBox "Error al Modificar (clsDESCUENTOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DESCUENTOS " & _
                "    WHERE ID_DESCUENTO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDESCUENTOS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DESCUENTOS ORDER BY NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_DESCUENTO) FROM DESCUENTOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_DESCUENTO = 1
    Else
        ID_DESCUENTO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_DESCUENTO, NOMBRE FROM DESCUENTOS " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
