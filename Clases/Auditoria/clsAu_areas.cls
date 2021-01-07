VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAu_areas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSAU_AREAS
'***************************************************************
Private ID_AREA As Integer
Private AREA As String
Private OBJETIVO As String
Private ALCANCE As String
Private FECHA_ALTA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSAU_AREAS
'***************************************************************
Public Property Let setID_AREA(ByVal dato As Integer)
        ID_AREA = dato
End Property
Public Property Let setAREA(ByVal dato As String)
        AREA = dato
End Property
Public Property Let setOBJETIVO(ByVal dato As String)
        OBJETIVO = dato
End Property
Public Property Let setALCANCE(ByVal dato As String)
        ALCANCE = dato
End Property
Public Property Let setFECHA_ALTA(ByVal dato As String)
        FECHA_ALTA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSAU_AREAS
'***************************************************************
Public Property Get getID_AREA() As Integer
        getID_AREA = ID_AREA
End Property
Public Property Get getAREA() As String
        getAREA = AREA
End Property
Public Property Get getOBJETIVO() As String
        getOBJETIVO = OBJETIVO
End Property
Public Property Get getALCANCE() As String
        getALCANCE = ALCANCE
End Property
Public Property Get getFECHA_ALTA() As String
        getFECHA_ALTA = FECHA_ALTA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSAU_AREAS
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM AU_AREAS WHERE ID_AREA = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_AREA = rs("ID_AREA")
                AREA = rs("AREA")
                OBJETIVO = rs("OBJETIVO")
                ALCANCE = rs("ALCANCE")
                FECHA_ALTA = rs("FECHA_ALTA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAu_areas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_AREA) FROM AU_AREAS"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_AREA = 1
        Else
                ID_AREA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO AU_AREAS " & _
                   "  VALUES (" & _
                        ID_AREA & "," & "'" & AREA & "'" & "," & "'" & OBJETIVO & "'" & "," & _
                        "'" & ALCANCE & "'" & "," & "'" & FECHA_ALTA & "'" & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_AREA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAu_areas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE AU_AREAS SET " & _
                        " AREA = '" & AREA & "'," & _
                        " OBJETIVO = '" & OBJETIVO & "'," & _
                        " ALCANCE = '" & ALCANCE & "'," & _
                        " FECHA_ALTA = '" & FECHA_ALTA & "'" & _
                " WHERE ID_AREA = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAu_areas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM AU_AREAS " & _
                "    WHERE ID_AREA = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAu_areas)", vbCritical, Err.Description
End Function
Public Function Listado(fAREA As String, fOBJETIVO As String, fALCANCE As String) As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM AU_AREAS " & _
                   " WHERE AREA LIKE '%" & fAREA & "%'" & _
                   "   AND OBJETIVO LIKE '%" & fOBJETIVO & "%'" & _
                   "   AND ALCANCE LIKE '%" & fALCANCE & "%'" & _
                   " ORDER BY AREA"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_AREA,AREA FROM AU_AREAS ORDER BY AREA"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "AU_AREAS"
        .setDESCRIPCION = "Áreas"
        .setPK = "ID_AREA"
        .setCAMPO = "AREA"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
