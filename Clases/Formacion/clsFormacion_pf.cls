VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_pf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_PLAN_FORMACION
'***************************************************************
Private ID_PLAN_FORMACION As Long
Private DESCRIPCION As String
Private OBSERVACIONES As String
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_PLAN_FORMACION
'***************************************************************
Public Property Let setID_PLAN_FORMACION(ByVal dato As Long)
        ID_PLAN_FORMACION = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property

Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_PLAN_FORMACION
'***************************************************************
Public Property Get getID_PLAN_FORMACION() As Long
        getID_PLAN_FORMACION = ID_PLAN_FORMACION
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_PLAN_FORMACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM FORMACION_PF WHERE ID_PLAN_FORMACION = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PLAN_FORMACION = rs("ID_PLAN_FORMACION")
                DESCRIPCION = rs("DESCRIPCION")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_plan_formacion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "INSERT INTO  FORMACION_PF" & _
                   "  VALUES (" & _
                        ID_PLAN_FORMACION & ",'" & DESCRIPCION & "','" & _
                        OBSERVACIONES & "'," & "CURRENT_TIMESTAMP" & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_PLAN_FORMACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_plan_formacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE FORMACION_PF SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " FTIMESTP = CURRENT_TIMESTAMP" & _
                " WHERE ID_PLAN_FORMACION = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_plan_formacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM  FORMACION_PF" & _
                "    WHERE ID_PLAN_FORMACION = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_plan_formacion)", vbCritical, Err.Description
End Function

Public Function Listado() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM FORMACION_PFORDER BY ID_PLAN_FORMACION"
        Set Listado = datos_bd(CONSULTA)
End Function

Public Function ListadoFiltro(cod As String, DESCRIPCION As String) As ADODB.Recordset
        Dim CONSULTA As String
        Dim strDescripcion As String
        
        If cod <> "" And IsNumeric(cod) Then
         CONSULTA = "SELECT * FROM FORMACION_PF WHERE ID_PLAN_FORMACION = " & CLng(cod)
        Else
             strDescripcion = " WHERE DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
             CONSULTA = "SELECT * FROM  FORMACION_PF" & _
                        strDescripcion
        End If
        Set ListadoFiltro = datos_bd(CONSULTA)
End Function

Public Sub CrearIdPlan()
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(ID_PLAN_FORMACION) FROM FORMACION_PF"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PLAN_FORMACION = 1
    Else
        ID_PLAN_FORMACION = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "FORMACION_PF"
        .setDESCRIPCION = "Planes de Formación"
        .setPK = "ID_NORMA"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_PLAN_FORMACION,DESCRIPCION FROM FORMACION_PF ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function


