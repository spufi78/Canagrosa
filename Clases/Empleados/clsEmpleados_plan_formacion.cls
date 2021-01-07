VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_plan_formacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEMPLEADOS_PLAN_FORMACION
'***************************************************************
Private ID_PLAN As Long
Private DESCRIPCION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEMPLEADOS_PLAN_FORMACION
'***************************************************************
Public Property Let setID_PLAN(ByVal dato As Long)
        ID_PLAN = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEMPLEADOS_PLAN_FORMACION
'***************************************************************
Public Property Get getID_PLAN() As Long
        getID_PLAN = ID_PLAN
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEMPLEADOS_PLAN_FORMACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM EMPLEADOS_PLAN_FORMACION WHERE ID_PLAN = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PLAN = rs("ID_PLAN")
                DESCRIPCION = rs("DESCRIPCION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpleados_plan_formacion)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_PLAN) FROM EMPLEADOS_PLAN_FORMACION"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PLAN = 1
        Else
                ID_PLAN = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO EMPLEADOS_PLAN_FORMACION " & _
                   "  VALUES (" & _
                        ID_PLAN & ",'" & DESCRIPCION & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_PLAN
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEmpleados_plan_formacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EMPLEADOS_PLAN_FORMACION SET " & _
                   " DESCRIPCION = '" & DESCRIPCION & "' " & _
                   " WHERE ID_PLAN = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEmpleados_plan_formacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EMPLEADOS_PLAN_FORMACION " & _
                   " WHERE ID_PLAN = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEmpleados_plan_formacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM EMPLEADOS_PLAN_FORMACION ORDER BY ID_PLAN"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_PLAN,DESCRIPCION FROM EMPLEADOS_PLAN_FORMACION ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
