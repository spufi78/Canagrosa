VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPresupuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPRESUPUESTOS
'***************************************************************
Private ID_PRESUPUESTO As Long
Private CLIENTE_ID As Long
Private fecha As String
Private DTO As Long
Private IVA As Long
Private TOTAL As Long
Private RECOGIDA As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPRESUPUESTOS
'***************************************************************
Public Property Let setID_PRESUPUESTO(ByVal dato As Long)
        ID_PRESUPUESTO = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setDTO(ByVal dato As Long)
        DTO = dato
End Property
Public Property Let setIVA(ByVal dato As Long)
        IVA = dato
End Property
Public Property Let setTOTAL(ByVal dato As Long)
        TOTAL = dato
End Property
Public Property Let setRECOGIDA(ByVal dato As Long)
        RECOGIDA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPRESUPUESTOS
'***************************************************************
Public Property Get getID_PRESUPUESTO() As Long
        getID_PRESUPUESTO = ID_PRESUPUESTO
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getDTO() As Long
        getDTO = DTO
End Property
Public Property Get getIVA() As Long
        getIVA = IVA
End Property
Public Property Get getTOTAL() As Long
        getTOTAL = TOTAL
End Property
Public Property Get getRECOGIDA() As Long
        getRECOGIDA = RECOGIDA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPRESUPUESTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PRESUPUESTOS WHERE ID_PRESUPUESTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PRESUPUESTO = rs("ID_PRESUPUESTO")
                CLIENTE_ID = rs("CLIENTE_ID")
                fecha = rs("FECHA")
                DTO = rs("DTO")
                IVA = rs("IVA")
                TOTAL = rs("TOTAL")
                RECOGIDA = rs("RECOGIDA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPresupuestos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_PRESUPUESTO) FROM PRESUPUESTOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PRESUPUESTO = 1
        Else
                ID_PRESUPUESTO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PRESUPUESTOS " & _
                   "  VALUES (" & _
                        ID_PRESUPUESTO & "," & CLIENTE_ID & "," & "'" & fecha & "'" & "," & _
                        DTO & "," & IVA & "," & TOTAL & "," & _
                        RECOGIDA & _
                ")"
        conn.Execute consulta
        Insertar = ID_PRESUPUESTO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPresupuestos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PRESUPUESTOS SET " & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " DTO = " & DTO & "," & _
                        " IVA = " & IVA & "," & _
                        " TOTAL = " & TOTAL & "," & _
                        " RECOGIDA = " & RECOGIDA & "" & _
                " WHERE ID_PRESUPUESTO = " & ID
        conn.Execute consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPresupuestos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PRESUPUESTOS_DETALLE WHERE PRESUPUESTO_ID = " & ID
        conn.Execute consulta
        consulta = "DELETE FROM PRESUPUESTOS WHERE ID_PRESUPUESTO = " & ID
        conn.Execute consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPresupuestos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_PRESUPUESTO,A.FECHA,C.NOMBRE,A.TOTAL " & _
                   "  FROM PRESUPUESTOS A, CLIENTES C " & _
                   " WHERE A.CLIENTE_ID = C.ID_CLIENTE " & _
                   " ORDER BY A.FECHA DESC "
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_ID(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_PRESUPUESTO,A.FECHA,C.NOMBRE,A.TOTAL " & _
                   "  FROM PRESUPUESTOS A, CLIENTES C " & _
                   " WHERE A.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND A.ID_PRESUPUESTO = " & ID & _
                   " ORDER BY A.FECHA DESC "
        Set Listado_por_ID = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PRESUPUESTO,CLIENTE_ID FROM PRESUPUESTOS ORDER BY CLIENTE_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
