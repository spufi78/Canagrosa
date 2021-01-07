VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_pf_docs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_PLAN_FORMACION_DOCS
'***************************************************************
Private PLAN_FORMACION_ID As Long
Private DOCUMENTO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_PLAN_FORMACION_DOCS
'***************************************************************
Public Property Let setPLAN_FORMACION_ID(ByVal dato As Long)
        PLAN_FORMACION_ID = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_PLAN_FORMACION_DOCS
'***************************************************************
Public Property Get getPLAN_FORMACION_ID() As Long
        getPLAN_FORMACION_ID = PLAN_FORMACION_ID
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_PLAN_FORMACION_DOCS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PF_DOCS WHERE PLAN_FORMACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PLAN_FORMACION_ID = rs("PLAN_FORMACION_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_pF_docs)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FORMACION_PF_DOCS " & _
                   "  VALUES (" & _
                        PLAN_FORMACION_ID & "," & DOCUMENTO_ID & _
                ")"
        execute_bd consulta
        Insertar = PLAN_FORMACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_plan_formacion_docs)", vbCritical, Err.Description
End Function
'M1106-I
Public Function Insertar_Ignorar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT IGNORE INTO FORMACION_PF_DOCS " & _
                   "  VALUES (" & _
                        PLAN_FORMACION_ID & "," & DOCUMENTO_ID & _
                ")"
        execute_bd consulta
        Insertar_Ignorar = PLAN_FORMACION_ID
        Exit Function
fallo:
        Insertar_Ignorar = 0
        MsgBox "Error al Insertar_Ignorar (clsFormacion_plan_formacion_docs)", vbCritical, Err.Description
End Function
'M1106-F
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_PF_DOCS SET " & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "" & _
                " WHERE PLAN_FORMACION_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_pF_docs)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_PF_DOCS " & _
                "    WHERE PLAN_FORMACION_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_pF_docs)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PF_DOCS ORDER BY PLAN_FORMACION_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Plan(planid As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PF_DOCS WHERE PLAN_FORMACION_ID = " & planid
        Set Listado_Plan = datos_bd(consulta)
End Function
