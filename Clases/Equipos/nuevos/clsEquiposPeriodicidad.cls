VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquiposPeriodicidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsEquiposPeriodicidad.cls
'***************************************************************
Private ID_PERIODICIDAD As Long
Private DESCRIPCION As String
Private DIAS As Long
Private MESES As Long
Private DIAS_PREAVISO As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsEquiposPeriodicidad.cls
'***************************************************************
Public Property Let setID_PERIODICIDAD(ByVal dato As Long)
        ID_PERIODICIDAD = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setDIAS(ByVal dato As Long)
        DIAS = dato
End Property
Public Property Let setMESES(ByVal dato As Long)
        MESES = dato
End Property
Public Property Let setDIAS_PREAVISO(ByVal dato As Long)
        DIAS_PREAVISO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsEquiposPeriodicidad.cls
'***************************************************************
Public Property Get getID_PERIODICIDAD() As Long
        getID_PERIODICIDAD = ID_PERIODICIDAD
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getDIAS() As Long
        getDIAS = DIAS
End Property
Public Property Get getMESES() As Long
        getMESES = MESES
End Property
Public Property Get getDIAS_PREAVISO() As Long
        getDIAS_PREAVISO = DIAS_PREAVISO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsEquiposPeriodicidad.cls
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM EQ_PERIODICIDAD WHERE ID_PERIODICIDAD = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PERIODICIDAD = rs("ID_PERIODICIDAD")
                DESCRIPCION = rs("DESCRIPCION")
                DIAS = rs("DIAS")
                MESES = rs("MESES")
                DIAS_PREAVISO = rs("DIAS_PREAVISO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquiposPeriodicidad.cls)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "INSERT INTO EQ_PERIODICIDAD " & _
                   "  VALUES (" & _
                        ID_PERIODICIDAD & "," & "'" & DESCRIPCION & "'" & "," & DIAS & "," & MESES & "," & _
                        DIAS_PREAVISO & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_PERIODICIDAD
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquiposPeriodicidad.cls)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE EQ_PERIODICIDAD SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " DIAS = " & DIAS & "," & _
                        " MESES = " & MESES & "," & _
                        " DIAS_PREAVISO = " & DIAS_PREAVISO & "" & _
                " WHERE ID_PERIODICIDAD = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquiposPeriodicidad.cls)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM EQ_PERIODICIDAD " & _
                "    WHERE ID_PERIODICIDAD = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquiposPeriodicidad.cls)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM EQ_PERIODICIDAD ORDER BY ID_PERIODICIDAD"
        Set Listado = datos_bd(CONSULTA)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_PERIODICIDAD, DESCRIPCION FROM EQ_PERIODICIDAD ORDER BY ID_PERIODICIDAD"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function

Public Function calcular_fecha(ByVal prm_fecha As Date, ByVal prm_id_periodo As Long, Optional ByVal avanza_retrasa As Integer = 1) As Date

    Dim DIAS As Integer
    If avanza_retrasa <> 1 And avanza_retrasa <> -1 Then
        avanza_retrasa = 1
    End If
        
    Dim oPer As New clsEquiposPeriodicidad
    oPer.Carga prm_id_periodo
    If oPer.getMESES <> 0 Then
        calcular_fecha = DateAdd("m", avanza_retrasa * oPer.getMESES, prm_fecha)
    Else
        calcular_fecha = DateAdd("d", avanza_retrasa * oPer.getDIAS, prm_fecha)
    End If
    Set oPer = Nothing
End Function

Public Function calcular_fecha_preaviso(ByVal prm_fecha As Date, ByVal prm_id_periodo As Long) As Date

    Dim DIAS As Integer
    Dim oPer As New clsEquiposPeriodicidad
    oPer.Carga prm_id_periodo
    
    DIAS = oPer.getDIAS_PREAVISO
    calcular_fecha_preaviso = DateAdd("d", -1 * DIAS, prm_fecha)
    
End Function

