VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_formacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEMPLEADOS_FORMACION
'***************************************************************
Private EMPLEADO_ID As Long
Private ID As Long
Private fecha As String
Private DESCRIPCION As String
Private ruta As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEMPLEADOS_FORMACION
'***************************************************************
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        ruta = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEMPLEADOS_FORMACION
'***************************************************************
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getRUTA() As String
        getRUTA = ruta
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEMPLEADOS_FORMACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM rrhh.empleados_formacion WHERE ID = " & ID
        Set rs = datos_bd_rrhh(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                EMPLEADO_ID = rs("EMPLEADO_ID")
                ID = rs("ID")
                fecha = rs("FECHA")
                DESCRIPCION = rs("DESCRIPCION")
                ruta = rs("RUTA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpleados_formacion)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM rrhh.empleados_formacion"
        Set rs = datos_bd_rrhh(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO rrhh.empleados_formacion " & _
                   "  VALUES (" & _
                        ID & "," & EMPLEADO_ID & ",'" & fecha & "','" & DESCRIPCION & "'" & "," & _
                        "'" & ruta & "'" & _
                ")"
        execute_bd_rrhh consulta
        Insertar = EMPLEADO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEmpleados_formacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE rrhh.empleados_formacion SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE ID = " & ID
        execute_bd_rrhh consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEmpleados_formacion)", vbCritical, Err.Description
End Function
Public Function ModificarAdjunto(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE rrhh.empleados_formacion " & _
                   "   SET RUTA = '" & ruta & "'" & _
                   " WHERE ID = " & ID
        execute_bd_rrhh consulta
        ModificarAdjunto = True
        Exit Function
fallo:
        ModificarAdjunto = False
        MsgBox "Error al ModificarAdjunto (clsEmpleados_formacion)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM rrhh.empleados_formacion " & _
                "    WHERE ID = " & ID
        execute_bd_rrhh consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEmpleados_formacion)", vbCritical, Err.Description
End Function
Public Function Listado(EMPLEADO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM rrhh.empleados_formacion WHERE EMPLEADO_ID = " & EMPLEADO & " ORDER BY FECHA DESC"
        Set Listado = datos_bd_rrhh(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT EMPLEADO_ID,ID FROM rrhh.empleados_formacion ORDER BY ID"
        Set Listado_Combo = datos_bd_rrhh(consulta)
End Function
