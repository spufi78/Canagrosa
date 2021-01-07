VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocs_pago_cobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDOCS_PAGO_COBROS
'***************************************************************
Private DOC_ID As Long
Private fecha As String
Private hora As String
Private EMPLEADO_ID As Long
Private FP_ID As Long
Private datos As String
Private OBSERVACIONES As String
Private fecha_prevista As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDOCS_PAGO_COBROS
'***************************************************************
Public Property Let setDOC_ID(ByVal dato As Long)
        DOC_ID = dato
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
Public Property Let setFECHA_PREVISTA(ByVal dato As String)
        fecha_prevista = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDOCS_PAGO_COBROS
'***************************************************************
Public Property Get getDOC_ID() As Long
        getDOC_ID = DOC_ID
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
Public Property Get getFECHA_PREVISTA() As String
        getFECHA_PREVISTA = fecha_prevista
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDOCS_PAGO_COBROS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCS_PAGO_COBROS WHERE DOC_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOC_ID = rs("DOC_ID")
                If Not IsNull(rs("FEcha")) Then
                    fecha = rs("FECHA")
                End If
                If Not IsNull(rs("hora")) Then
                    hora = rs("HORA")
                End If
                EMPLEADO_ID = rs("EMPLEADO_ID")
                FP_ID = rs("FP_ID")
                datos = rs("DATOS")
                OBSERVACIONES = rs("OBSERVACIONES")
                If Not IsNull(rs("FECHA_PREVISTA")) Then
                    fecha_prevista = rs("FECHA_PREVISTA")
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDocs_pago_cobros)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE INTO DOCS_PAGO_COBROS " & _
                   "  VALUES (" & _
                        DOC_ID & "," & fecha & ",'" & hora & "'," & _
                        EMPLEADO_ID & "," & FP_ID & "," & "'" & datos & "'" & "," & _
                        "'" & OBSERVACIONES & "'," & fecha_prevista & _
                ")"
        execute_bd consulta
        Insertar = DOC_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDocs_pago_cobros)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCS_PAGO_COBROS SET " & _
                        " FECHA = " & fecha & "," & _
                        " FECHA_PREVISTA = " & fecha_prevista & "," & _
                        " HORA = " & hora & "," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " FP_ID = " & FP_ID & "," & _
                        " DATOS = '" & datos & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE DOC_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDocs_pago_cobros)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCS_PAGO_COBROS " & _
                "    WHERE DOC_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDocs_pago_cobros)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCS_PAGO_COBROS ORDER BY DOC_ID"
        Set Listado = datos_bd(consulta)
End Function
