VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVehiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSVEHICULOS
'***************************************************************
Private ID_VEHICULO As Long
Private NOMBRE As String
Private MATRICULA As String
Private Nif As String
Private REMOLQUE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSVEHICULOS
'***************************************************************
Public Property Let setID_VEHICULO(ByVal dato As Long)
        ID_VEHICULO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setMATRICULA(ByVal dato As String)
        MATRICULA = dato
End Property
Public Property Let setNIF(ByVal dato As String)
        Nif = dato
End Property
Public Property Let setREMOLQUE(ByVal dato As String)
        REMOLQUE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSVEHICULOS
'***************************************************************
Public Property Get getID_VEHICULO() As Long
        getID_VEHICULO = ID_VEHICULO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getMATRICULA() As String
        getMATRICULA = MATRICULA
End Property
Public Property Get getNIF() As String
        getNIF = Nif
End Property
Public Property Get getREMOLQUE() As String
        getREMOLQUE = REMOLQUE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSVEHICULOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM VEHICULOS WHERE ID_VEHICULO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_VEHICULO = rs("ID_VEHICULO")
                NOMBRE = rs("NOMBRE")
                MATRICULA = rs("MATRICULA")
                Nif = rs("NIF")
                REMOLQUE = rs("REMOLQUE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsVehiculos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_VEHICULO) FROM VEHICULOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_VEHICULO = 1
        Else
                ID_VEHICULO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_VEHICULO = 0 Then
            Me.CrearID
        End If
        consulta = "INSERT INTO VEHICULOS " & _
                   "  VALUES (" & _
                        ID_VEHICULO & "," & "'" & NOMBRE & "'" & "," & "'" & MATRICULA & "'" & "," & _
                        "'" & Nif & "'" & "," & "'" & REMOLQUE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_VEHICULO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsVehiculos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE VEHICULOS SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " MATRICULA = '" & MATRICULA & "'," & _
                        " NIF = '" & Nif & "'," & _
                        " REMOLQUE = '" & REMOLQUE & "'" & _
                " WHERE ID_VEHICULO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsVehiculos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT * FROM DOCUMENTOS WHERE VEHICULO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar el vehículo. Tiene albaranes asignados.", vbExclamation, App.Title
            Eliminar = False
            Exit Function
        End If
        consulta = "DELETE FROM VEHICULOS " & _
                "    WHERE ID_VEHICULO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsVehiculos)", vbCritical, Err.Description
End Function
Public Function Listado(NOMBRE As String, MATRICULA As String) As ADODB.Recordset
        Dim consulta As String
        Dim s As String
        If NOMBRE <> "" Then
            s = s & "  AND NOMBRE LIKE '%" & NOMBRE & "%'"
        End If
        If MATRICULA <> "" Then
            s = s & "  AND MATRICULA LIKE '%" & MATRICULA & "%'"
        End If
        consulta = "SELECT * FROM VEHICULOS " & _
                   " WHERE 1 = 1 " & s & _
                   " ORDER BY NOMBRE,MATRICULA,REMOLQUE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_VEHICULO,NOMBRE FROM VEHICULOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(combo As miComboBCA, FK As Long, FORMULARIO As Form, FILTRO As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "VEHICULOS"
        .setDESCRIPCION = "Vehiculos"
        .setPK = "ID_VEHICULO"
        .setCAMPO = "concat(CAST(ID_VEHICULO AS CHAR),' - [',MATRICULA,'] [',REMOLQUE,'] ', NOMBRE, ' (',NIF,')')"
        .setMUESTRA_DETALLE = True
        .setFILTRO = FILTRO
        Set .FORMULARIO = FORMULARIO
    End With
End Function

