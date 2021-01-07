VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTAREAS
'***************************************************************
Private ID_TAREA As Long
Private MODULO_ID As Long
Private DESCRIPCION As String
Private FECHA_ALTA As String
Private FECHA_BAJA As String
Private ACTIVA As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTAREAS
'***************************************************************
Public Property Let setID_TAREA(ByVal dato As Long)
        ID_TAREA = dato
End Property
Public Property Let setMODULO_ID(ByVal dato As Long)
        MODULO_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFECHA_ALTA(ByVal dato As String)
        FECHA_ALTA = dato
End Property
Public Property Let setFECHA_BAJA(ByVal dato As String)
        FECHA_BAJA = dato
End Property
Public Property Let setACTIVA(ByVal dato As Integer)
        ACTIVA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTAREAS
'***************************************************************
Public Property Get getID_TAREA() As Long
        getID_TAREA = ID_TAREA
End Property
Public Property Get getMODULO_ID() As Long
        getMODULO_ID = MODULO_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA_ALTA() As String
        getFECHA_ALTA = FECHA_ALTA
End Property
Public Property Get getFECHA_BAJA() As String
        getFECHA_BAJA = FECHA_BAJA
End Property
Public Property Get getACTIVA() As Integer
        getACTIVA = ACTIVA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTAREAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM tareas WHERE ID_TAREA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TAREA = rs("ID_TAREA")
                MODULO_ID = rs("MODULO_ID")
                DESCRIPCION = rs("DESCRIPCION")
                FECHA_ALTA = rs("FECHA_ALTA")
                FECHA_BAJA = rs("FECHA_BAJA")
                ACTIVA = rs("ACTIVA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTareas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_TAREA) FROM TAREAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_TAREA = 1
        Else
                ID_TAREA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO tareas " & _
                   "  VALUES (" & _
                        ID_TAREA & "," & MODULO_ID & "," & "'" & DESCRIPCION & "'" & "," & _
                        "'" & FECHA_ALTA & "'" & "," & "'" & FECHA_BAJA & "'" & "," & ACTIVA & _
                ")"
        execute_bd consulta
        Insertar = ID_TAREA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTareas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE tareas SET " & _
                        " MODULO_ID = " & MODULO_ID & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FECHA_ALTA = '" & FECHA_ALTA & "'," & _
                        " FECHA_BAJA = '" & FECHA_BAJA & "'," & _
                        " ACTIVA = " & ACTIVA & "" & _
                " WHERE ID_TAREA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTareas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("SELECT * FROM tareas_incurridos WHERE TAREA_ID = " & ID)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar la tarea. Tiene horas incurridas.", vbExclamation, App.Title
            Eliminar = False
        Else
            execute_bd "DELETE FROM tareas WHERE ID_TAREA = " & ID
            Eliminar = True
        End If
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTareas)", vbCritical, Err.Description
End Function
Public Function Listado(MODULO As Long, DESCRIPCION As String, ACTIVA As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim s As String
        If MODULO <> 0 Then
            s = " AND A.MODULO_ID = " & MODULO
        End If
        If ACTIVA = True Then
            s = s & " AND A.ACTIVA = 1"
        End If
        consulta = "SELECT A.ID_TAREA,B.DESCRIPCION,A.DESCRIPCION,A.FECHA_ALTA,A.FECHA_BAJA,A.ACTIVA " & _
                   "  FROM tareas A, decodificadora B " & _
                   " WHERE B.CODIGO = " & DECODIFICADORA.TAREAS_MODULOS & _
                   "   AND B.VALOR  = A.MODULO_ID " & _
                   "   AND A.DESCRIPCION LIKE '%" & DESCRIPCION & "%' " & _
                   s & _
                   " ORDER BY B.DESCRIPCION,A.DESCRIPCION "
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_Codigo(ID_TAREA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_TAREA,B.DESCRIPCION,A.DESCRIPCION,A.FECHA_ALTA,A.FECHA_BAJA,A.ACTIVA " & _
                   "  FROM tareas A, decodificadora B " & _
                   " WHERE B.CODIGO = " & DECODIFICADORA.TAREAS_MODULOS & _
                   "   AND B.VALOR  = A.MODULO_ID " & _
                   "   AND A.ID_TAREA = " & ID_TAREA
        Set Listado_por_Codigo = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_TAREA,DESCRIPCION FROM tareas ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_Combo_Filtro(MODULO As Long, fecha As Date) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_TAREA,DESCRIPCION " & _
                   "  FROM tareas " & _
                   " WHERE MODULO_ID = " & MODULO & _
                   "   AND ACTIVA = 1 " & _
                   "   AND FECHA_ALTA <= '" & Format(fecha, "yyyy-mm-dd") & "' " & _
                   "   AND FECHA_BAJA >= '" & Format(fecha, "yyyy-mm-dd") & "' " & _
                   " ORDER BY DESCRIPCION"
        Set Listado_Combo_Filtro = datos_bd(consulta)
End Function

