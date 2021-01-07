VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedores_riesgos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROVEEDORES_RIESGOS
'***************************************************************
Private ID As Long
Private PROVEEDOR_ID As Long
Private fecha As String
Private USUARIO_ID As Long
Private tipo As Long
Private VALOR As String
Private PROBABILIDAD As Long
Private PLAN As String
Private TAREA As String
Private RIESGO As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROVEEDORES_RIESGOS
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
Public Property Let setVALOR(ByVal dato As Long)
        VALOR = dato
End Property
Public Property Let setPROBABILIDAD(ByVal dato As Long)
        PROBABILIDAD = dato
End Property
Public Property Let setPLAN(ByVal dato As String)
        PLAN = dato
End Property
Public Property Let setTAREA(ByVal dato As String)
        TAREA = dato
End Property
Public Property Let setRIESGO(ByVal dato As Long)
        RIESGO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROVEEDORES_RIESGOS
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTIPO() As Integer
        getTIPO = tipo
End Property
Public Property Get getVALOR() As Integer
        getVALOR = VALOR
End Property
Public Property Get getPROBABILIDAD() As Long
        getPROBABILIDAD = PROBABILIDAD
End Property
Public Property Get getPLAN() As String
        getPLAN = PLAN
End Property
Public Property Get getTAREA() As String
        getTAREA = TAREA
End Property
Public Property Get getRIESGO() As Long
        getRIESGO = RIESGO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROVEEDORES_RIESGOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_RIESGOS WHERE ID = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                ID = RS("ID")
                PROVEEDOR_ID = RS("PROVEEDOR_ID")
                fecha = RS("FECHA")
                USUARIO_ID = RS("USUARIO_ID")
                tipo = RS("TIPO")
                VALOR = RS("VALOR")
                PROBABILIDAD = RS("PROBABILIDAD")
                PLAN = RS("PLAN")
                TAREA = RS("TAREA")
                RIESGO = RS("RIESGO")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsProveedores_riesgos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM PROVEEDORES_RIESGOS"
        Set RS = datos_bd(consulta)
        If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then
                ID = 1
        Else
                ID = RS.Fields(0) + 1
        End If
        Set RS = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PROVEEDORES_RIESGOS " & _
                   "  VALUES (" & _
                        ID & "," & PROVEEDOR_ID & ",CURRENT_DATE," & _
                        USUARIO.getID_EMPLEADO & "," & tipo & "," & VALOR & "," & _
                        PROBABILIDAD & ",'" & PLAN & "','" & TAREA & "'," & RIESGO & _
                ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProveedores_riesgos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDORES_RIESGOS SET " & _
                        " FECHA = CURRENT_DATE," & _
                        " PLAN = '" & PLAN & "'," & _
                        " TAREA = '" & TAREA & "'," & _
                        " USUARIO_ID = " & USUARIO.getID_EMPLEADO & "," & _
                        " TIPO = " & tipo & "," & _
                        " VALOR = " & VALOR & "," & _
                        " PROBABILIDAD = " & PROBABILIDAD & "," & _
                        " RIESGO = " & RIESGO & "" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsProveedores_riesgos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROVEEDORES_RIESGOS " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsProveedores_riesgos)", vbCritical, Err.Description
End Function
Public Function Listado(proveedorId As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID, A.FECHA, B.USUARIO, A.TIPO, A.VALOR, A.PROBABILIDAD, A.RIESGO, A.PLAN, A.TAREA " & _
                   "  FROM PROVEEDORES_RIESGOS A " & _
                   "  LEFT JOIN usuarios B ON A.USUARIO_ID = B.ID_EMPLEADO " & _
                   " WHERE A.PROVEEDOR_ID = " & proveedorId & _
                   " ORDER BY A.FECHA DESC "
        Set Listado = datos_bd(consulta)
End Function
