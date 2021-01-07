VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedores_evaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROVEEDORES_EVALUACION
'***************************************************************
Private ID_EVALUACION As Long
Private PROVEEDOR_ID As Long
Private fecha As String
Private DESCRIPCION As String
Private OBSERVACIONES As String
Private USUARIO_ID As Long
Private TS As Long
Private OP1 As Long
Private OP2 As Long
Private OP3 As Long
Private OP4 As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROVEEDORES_EVALUACION
'***************************************************************
Public Property Let setID_EVALUACION(ByVal dato As Long)
        ID_EVALUACION = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTS(ByVal dato As Long)
        TS = dato
End Property
Public Property Let setOP1(ByVal dato As Long)
        OP1 = dato
End Property
Public Property Let setOP2(ByVal dato As Long)
        OP2 = dato
End Property
Public Property Let setOP3(ByVal dato As Long)
        OP3 = dato
End Property
Public Property Let setOP4(ByVal dato As Long)
        OP4 = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROVEEDORES_EVALUACION
'***************************************************************
Public Property Get getID_EVALUACION() As Long
        getID_EVALUACION = ID_EVALUACION
End Property
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTS() As Long
        getTS = TS
End Property
Public Property Get getOP1() As Long
        getOP1 = OP1
End Property
Public Property Get getOP2() As Long
        getOP2 = OP2
End Property
Public Property Get getOP3() As Long
        getOP3 = OP3
End Property
Public Property Get getOP4() As Long
        getOP4 = OP4
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROVEEDORES_EVALUACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_EVALUACION WHERE ID_EVALUACION = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                ID_EVALUACION = RS("ID_EVALUACION")
                PROVEEDOR_ID = RS("PROVEEDOR_ID")
                fecha = RS("FECHA")
                DESCRIPCION = RS("DESCRIPCION")
                OBSERVACIONES = RS("OBSERVACIONES")
                USUARIO_ID = RS("USUARIO_ID")
                TS = RS("TS")
                OP1 = RS("OP1")
                OP2 = RS("OP2")
                OP3 = RS("OP3")
                OP4 = RS("OP4")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsProveedores_evaluacion)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_EVALUACION) FROM PROVEEDORES_EVALUACION"
        Set RS = datos_bd(consulta)
        If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then
                ID_EVALUACION = 1
        Else
                ID_EVALUACION = RS.Fields(0) + 1
        End If
        Set RS = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PROVEEDORES_EVALUACION " & _
                   "  VALUES (" & _
                        ID_EVALUACION & "," & PROVEEDOR_ID & "," & "'" & fecha & "'" & "," & _
                        "'" & DESCRIPCION & "'" & "," & "'" & OBSERVACIONES & "'" & "," & USUARIO_ID & "," & _
                        TS & "," & OP1 & "," & OP2 & "," & _
                        OP3 & "," & OP4 & _
                ")"
        execute_bd consulta
        Insertar = ID_EVALUACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProveedores_evaluacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDORES_EVALUACION SET " & _
                        " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TS = " & TS & "," & _
                        " OP1 = " & OP1 & "," & _
                        " OP2 = " & OP2 & "," & _
                        " OP3 = " & OP3 & "," & _
                        " OP4 = " & OP4 & "" & _
                " WHERE ID_EVALUACION = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsProveedores_evaluacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROVEEDORES_EVALUACION " & _
                "    WHERE ID_EVALUACION = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsProveedores_evaluacion)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_EVALUACION,A.FECHA,A.DESCRIPCION,A.OBSERVACIONES,A.OP1,A.OP2,A.OP3,A.OP4,B.USUARIO " & _
                   "  FROM PROVEEDORES_EVALUACION A,usuarios B " & _
                   " WHERE A.USUARIO_ID = B.ID_EMPLEADO " & _
                   "   AND A.PROVEEDOR_ID = " & ID & _
                   " ORDER BY A.FECHA DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_EVALUACION,PROVEEDOR_ID FROM PROVEEDORES_EVALUACION ORDER BY PROVEEDOR_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
