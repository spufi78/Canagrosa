VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_cualificaciones_m"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES_MUESTRAS
'***************************************************************
Private CUALIFICACION_ID As Long
Private MUESTRA_ID As Long
Private tipo As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES_MUESTRAS
'***************************************************************
Public Property Let setCUALIFICACION_ID(ByVal dato As Long)
        CUALIFICACION_ID = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES_MUESTRAS
'***************************************************************
Public Property Get getCUALIFICACION_ID() As Long
        getCUALIFICACION_ID = CUALIFICACION_ID
End Property
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES_MUESTRAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM empleados_cualificaciones_muestras WHERE CUALIFICACION_ID = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                CUALIFICACION_ID = RS("CUALIFICACION_ID")
                MUESTRA_ID = RS("MUESTRA_ID")
                tipo = RS("TIPO")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpleados_cualificaciones_muestras)", vbCritical, Err.Description
End Function
Public Function CargaPorMuestra(MUESTRA_ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM empleados_cualificaciones_muestras WHERE MUESTRA_ID = " & MUESTRA_ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                CargaPorMuestra = False
        Else
                CUALIFICACION_ID = RS("CUALIFICACION_ID")
                MUESTRA_ID = RS("MUESTRA_ID")
                tipo = RS("TIPO")
                CargaPorMuestra = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        CargaPorMuestra = False
        MsgBox "Error al cargar los datos(clsEmpleados_cualificaciones_muestras)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO empleados_cualificaciones_muestras " & _
                   "  VALUES (" & _
                        CUALIFICACION_ID & "," & MUESTRA_ID & "," & tipo & _
                ")"
        execute_bd consulta
        Insertar = CUALIFICACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEmpleados_cualificaciones_muestras)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE empleados_cualificaciones_muestras SET " & _
                        " MUESTRA_ID = " & MUESTRA_ID & "," & _
                        " TIPO = " & tipo & "" & _
                " WHERE CUALIFICACION_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEmpleados_cualificaciones_muestras)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM empleados_cualificaciones_muestras " & _
                "    WHERE CUALIFICACION_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEmpleados_cualificaciones_muestras)", vbCritical, Err.Description
End Function
Public Function EliminarIndividual(ID As Long, MUESTRA_ID As Long, TIPO_ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM empleados_cualificaciones_muestras " & _
                "    WHERE CUALIFICACION_ID = " & ID & _
                "    AND MUESTRA_ID = " & MUESTRA_ID & _
                "    AND TIPO = " & tipo
        execute_bd consulta
        EliminarIndividual = True
        Exit Function
fallo:
        EliminarIndividual = False
        MsgBox "Error al EliminarIndividual (clsEmpleados_cualificaciones_muestras)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ECM.TIPO, ECM.MUESTRA_ID, CONCAT(TM.codigo,'-',CAST(ID_PARTICULAR AS CHAR)) AS CODIGO, ID_GENERAL " & _
                   "  FROM empleados_cualificaciones_muestras ECM, muestras M, tipos_muestra TM " & _
                   " WHERE M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
                   "   AND ECM.MUESTRA_ID = M.ID_MUESTRA " & _
                   "   AND ECM.CUALIFICACION_ID = " & ID & _
                   " ORDER BY ECM.MUESTRA_ID ASC"
        Set Listado = datos_bd(consulta)
End Function
