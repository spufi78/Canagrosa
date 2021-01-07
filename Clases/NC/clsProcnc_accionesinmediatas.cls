VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcnc_accionesinmediatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROCNC_ACCIONESINMEDIATAS
'***************************************************************
Private id_accion_inmediata As Long
Private id_procnc As Long
Private descripcion As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROCNC_ACCIONESINMEDIATAS
'***************************************************************
Public Property Let setID_ACCION_INMEDIATA(ByVal dato As Long)
        id_accion_inmediata = dato
End Property
Public Property Let setID_PROCNC(ByVal dato As Long)
        id_procnc = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        descripcion = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROCNC_ACCIONESINMEDIATAS
'***************************************************************
Public Property Get getID_ACCION_INMEDIATA() As Long
        getID_ACCION_INMEDIATA = id_accion_inmediata
End Property
Public Property Get getID_PROCNC() As Long
        getID_PROCNC = id_procnc
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = descripcion
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROCNC_ACCIONESINMEDIATAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM PROCNC_ACCIONESINMEDIATAS WHERE id_accion_inmediata = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                id_accion_inmediata = rs("id_accion_inmediata")
                id_procnc = rs("id_procnc")
                descripcion = rs("descripcion")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsProcnc_accionesinmediatas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(id_accion_inmediata) FROM PROCNC_ACCIONESINMEDIATAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                id_accion_inmediata = 1
        Else
                id_accion_inmediata = CLng(rs.Fields(0)) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PROCNC_ACCIONESINMEDIATAS " & _
                   "  VALUES (" & _
                        id_accion_inmediata & "," & id_procnc & "," & "'" & descripcion & "'" & _
                ")"
        execute_bd consulta
        Insertar = id_accion_inmediata
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProcnc_accionesinmediatas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROCNC_ACCIONESINMEDIATAS SET " & _
                        " ID_PROCNC = " & id_procnc & "," & _
                        " DESCRIPCION = '" & descripcion & "'" & _
                " WHERE id_accion_inmediata = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsProcnc_accionesinmediatas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROCNC_ACCIONESINMEDIATAS " & _
                "    WHERE ID_PROCNC = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsProcnc_accionesinmediatas)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM PROCNC_ACCIONESINMEDIATAS " & _
                   " WHERE ID_PROCNC = " & ID & _
                   " ORDER BY id_accion_inmediata"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_ACCION_INMEDIATA,ID_PROCNC FROM PROCNC_ACCIONESINMEDIATAS ORDER BY id_procnc"
        Set Listado_Combo = datos_bd(consulta)
End Function
