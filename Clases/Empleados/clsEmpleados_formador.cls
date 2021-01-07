VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_formador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEMPLEADOS_FORMADOR
'***************************************************************
Private EMPLEADO_ID_FORMADOR As Long
Private EMPLEADO_ID_FORMADO As Long
Private DOCUMENTO_ID As Long
Private fecha As String
Private ADJUNTOS As String
Private OBSERVACIONES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEMPLEADOS_FORMADOR
'***************************************************************
Public Property Let setEMPLEADO_ID_FORMADOR(ByVal dato As Long)
        EMPLEADO_ID_FORMADOR = dato
End Property
Public Property Let setEMPLEADO_ID_FORMADO(ByVal dato As Long)
        EMPLEADO_ID_FORMADO = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setADJUNTOS(ByVal dato As String)
        ADJUNTOS = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEMPLEADOS_FORMADOR
'***************************************************************
Public Property Get getEMPLEADO_ID_FORMADOR() As Long
        getEMPLEADO_ID_FORMADOR = EMPLEADO_ID_FORMADOR
End Property
Public Property Get getEMPLEADO_ID_FORMADO() As Long
        getEMPLEADO_ID_FORMADO = EMPLEADO_ID_FORMADO
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getADJUNTOS() As String
        getADJUNTOS = ADJUNTOS
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEMPLEADOS_FORMADOR
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EMPLEADOS_FORMADOR WHERE EMPLEADO_ID_FORMADOR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                EMPLEADO_ID_FORMADOR = rs("EMPLEADO_ID_FORMADOR")
                EMPLEADO_ID_FORMADO = rs("EMPLEADO_ID_FORMADO")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                fecha = rs("FECHA")
                ADJUNTOS = rs("ADJUNTOS")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpleados_formador)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO EMPLEADOS_FORMADOR " & _
                   "  VALUES (" & _
                        EMPLEADO_ID_FORMADOR & "," & EMPLEADO_ID_FORMADO & "," & DOCUMENTO_ID & "," & _
                        "'" & fecha & "'" & "," & "'" & ADJUNTOS & "'" & "," & "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = EMPLEADO_ID_FORMADOR
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEmpleados_formador)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EMPLEADOS_FORMADOR SET " & _
                        " EMPLEADO_ID_FORMADO = " & EMPLEADO_ID_FORMADO & "," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " ADJUNTOS = '" & ADJUNTOS & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE EMPLEADO_ID_FORMADOR = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEmpleados_formador)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EMPLEADOS_FORMADOR " & _
                "    WHERE EMPLEADO_ID_FORMADOR = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEmpleados_formador)", vbCritical, Err.Description
End Function
Public Function Listado(ID_EMPLEADO As Long, PNT As String, formado As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT EF.EMPLEADO_ID_FORMADO, EF.DOCUMENTO_ID, CD.NOMBRE, EF.FECHA, " & _
                   "       E.NOMBRE " & _
                   "  FROM EMPLEADOS_FORMADOR EF " & _
                   "  LEFT JOIN empleados E ON EF.EMPLEADO_ID_FORMADO = E.ID_EMPLEADO " & _
                   "  LEFT JOIN ca_documentos CD ON EF.DOCUMENTO_ID = CD.ID_DOCUMENTO " & _
                   " WHERE EF.EMPLEADO_ID_FORMADOR = " & ID_EMPLEADO & _
                   "   AND CD.NOMBRE LIKE '%" & PNT & "%'" & _
                   IIf(formado <> 0, " AND EF.EMPLEADO_ID_FORMADO = " & formado, "") & _
                   " ORDER BY FECHA DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_ID(xEMPLEADO_ID_FORMADOR As Long, xEMPLEADO_ID_FORMADO As Long, xDOCUMENTO_ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT EF.EMPLEADO_ID_FORMADO, EF.DOCUMENTO_ID, CD.NOMBRE, EF.FECHA, " & _
                   "       E.NOMBRE " & _
                   "  FROM EMPLEADOS_FORMADOR EF " & _
                   "  LEFT JOIN empleados E ON EF.EMPLEADO_ID_FORMADO = E.ID_EMPLEADO " & _
                   "  LEFT JOIN ca_documentos CD ON EF.DOCUMENTO_ID = CD.ID_DOCUMENTO " & _
                   " WHERE EF.EMPLEADO_ID_FORMADOR = " & xEMPLEADO_ID_FORMADOR & _
                   "   AND EF.EMPLEADO_ID_FORMADO = " & xEMPLEADO_ID_FORMADO & _
                   "   AND EF.DOCUMENTO_ID = " & xDOCUMENTO_ID
        Set Listado_ID = datos_bd(consulta)
End Function

