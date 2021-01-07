VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentos_cobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsDocumentos_cobros
'***************************************************************
Private DOCUMENTO_ID As Long
Private VENCIMIENTO As Long
Private fecha As String
Private hora As String
Private EMPLEADO_ID As Long
Private FP_ID As Long
Private datos As String
Private OBSERVACIONES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsDocumentos_cobros
'***************************************************************
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setVENCIMIENTO(ByVal dato As Long)
        VENCIMIENTO = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setHORA(ByVal dato As String)
        hora = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setFP_ID(ByVal dato As Long)
        FP_ID = dato
End Property
Public Property Let setDATOS(ByVal dato As String)
        datos = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsDocumentos_cobros
'***************************************************************
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getVENCIMIENTO() As Long
        getVENCIMIENTO = VENCIMIENTO
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getHORA() As String
        getHORA = hora
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getFP_ID() As Long
        getFP_ID = FP_ID
End Property
Public Property Get getDATOS() As String
        getDATOS = datos
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsDocumentos_cobros
'***************************************************************
Public Function Carga(ID As Long, VENCIMIENTO As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_COBROS WHERE DOCUMENTO_ID = " & ID & " AND VENCIMIENTO = " & VENCIMIENTO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                VENCIMIENTO = rs("VENCIMIENTO")
                fecha = rs("FECHA")
                hora = rs("HORA")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                FP_ID = rs("FP_ID")
                datos = rs("DATOS")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDocumentos_cobros)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO DOCUMENTOS_COBROS " & _
                   "  VALUES (" & _
                        DOCUMENTO_ID & "," & VENCIMIENTO & ",'" & fecha & "'" & ",'" & hora & "'," & _
                        EMPLEADO_ID & "," & FP_ID & "," & "'" & datos & "'" & "," & _
                        "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = DOCUMENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDocumentos_cobros)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS_COBROS SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " HORA = " & hora & "," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " FP_ID = " & FP_ID & "," & _
                        " DATOS = '" & datos & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDocumentos_cobros)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCUMENTOS_COBROS " & _
                "    WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDocumentos_cobros)", vbCritical, Err.Description
End Function
Public Function Eliminar_Vencimiento(ID As Long, VENCIMIENTO As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCUMENTOS_COBROS " & _
                "    WHERE DOCUMENTO_ID = " & ID & " AND VENCIMIENTO = " & VENCIMIENTO
        execute_bd consulta
        Eliminar_Vencimiento = True
        Exit Function
fallo:
        Eliminar_Vencimiento = False
        MsgBox "Error al Eliminar_Vencimiento (clsDocumentos_cobros)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_COBROS ORDER BY DOCUMENTO_ID"
        Set Listado = datos_bd(consulta)
End Function
