VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuministros_tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSUMINISTROS_TIPOS
'***************************************************************
Private ID_TIPO_SUMINISTRO As Long
Private producto As String
Private PROCEDIMIENTO As String
Private ID_REACTIVO_PR As Long
Private TIPO_CADUCIDAD_ID As Long
Private anulado As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSUMINISTROS_TIPOS
'***************************************************************
Public Property Let setID_TIPO_SUMINISTRO(ByVal dato As Long)
        ID_TIPO_SUMINISTRO = dato
End Property
Public Property Let setPRODUCTO(ByVal dato As String)
        producto = dato
End Property
Public Property Let setPROCEDIMIENTO(ByVal dato As String)
        PROCEDIMIENTO = dato
End Property
Public Property Let setID_REACTIVO_PR(ByVal dato As Long)
        ID_REACTIVO_PR = dato
End Property
Public Property Let setTIPO_CADUCIDAD_ID(ByVal dato As Long)
        TIPO_CADUCIDAD_ID = dato
End Property
Public Property Let setANULADO(ByVal dato As Long)
        anulado = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSUMINISTROS_TIPOS
'***************************************************************
Public Property Get getID_TIPO_SUMINISTRO() As Long
        getID_TIPO_SUMINISTRO = ID_TIPO_SUMINISTRO
End Property
Public Property Get getPRODUCTO() As String
        getPRODUCTO = producto
End Property
Public Property Get getPROCEDIMIENTO() As String
        getPROCEDIMIENTO = PROCEDIMIENTO
End Property
Public Property Get getID_REACTIVO_PR() As Long
        getID_REACTIVO_PR = ID_REACTIVO_PR
End Property
Public Property Get getTIPO_CADUCIDAD_ID() As Long
        getTIPO_CADUCIDAD_ID = TIPO_CADUCIDAD_ID
End Property
Public Property Get getANULADO() As Long
        getANULADO = anulado
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSUMINISTROS_TIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM SUMINISTROS_TIPOS WHERE ID_TIPO_SUMINISTRO = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TIPO_SUMINISTRO = rs("ID_TIPO_SUMINISTRO")
                producto = rs("PRODUCTO")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                ID_REACTIVO_PR = rs("ID_REACTIVO_PR")
                TIPO_CADUCIDAD_ID = rs("TIPO_CADUCIDAD_ID")
                anulado = rs("ANULADO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSuministros_tipos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_TIPO_SUMINISTRO) FROM SUMINISTROS_TIPOS"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_TIPO_SUMINISTRO = 1
        Else
                ID_TIPO_SUMINISTRO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO SUMINISTROS_TIPOS " & _
                   "  VALUES (" & _
                        ID_TIPO_SUMINISTRO & "," & "'" & producto & "'" & "," & "'" & PROCEDIMIENTO & "'" & "," & _
                        ID_REACTIVO_PR & "," & TIPO_CADUCIDAD_ID & ",0" & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_TIPO_SUMINISTRO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSuministros_tipos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE SUMINISTROS_TIPOS SET " & _
                        " PRODUCTO = '" & producto & "'," & _
                        " PROCEDIMIENTO = '" & PROCEDIMIENTO & "'," & _
                        " TIPO_CADUCIDAD_ID = " & TIPO_CADUCIDAD_ID & "," & _
                        " ID_REACTIVO_PR = " & ID_REACTIVO_PR & "" & _
                " WHERE ID_TIPO_SUMINISTRO = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSuministros_tipos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE SUMINISTROS_TIPOS " & _
                   "   SET ANULADO = " & usuario.getID_EMPLEADO & _
                   "  WHERE ID_TIPO_SUMINISTRO = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSuministros_tipos)", vbCritical, Err.Description
End Function
Public Function Reactivar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE SUMINISTROS_TIPOS " & _
                   "   SET ANULADO = 0 " & _
                   "  WHERE ID_TIPO_SUMINISTRO = " & ID
        execute_bd CONSULTA
        Reactivar = True
        Exit Function
fallo:
        Reactivar = False
        MsgBox "Error al Reactivar (clsSuministros_tipos)", vbCritical, Err.Description
End Function

Public Function Listado(producto As String, PROCEDIMIENTO As String, ANULADOS As Boolean) As ADODB.Recordset
        Dim CONSULTA As String
        Dim filtro As String
        If producto <> "" Then
            filtro = filtro & " and A.PRODUCTO LIKE '%" & producto & "%'"
        End If
        If PROCEDIMIENTO <> "" Then
            filtro = filtro & " and A.PROCEDIMIENTO LIKE '%" & PROCEDIMIENTO & "%'"
        End If
        If ANULADOS = False Then
            filtro = filtro & " and A.ANULADO = 0 "
        End If
        CONSULTA = "SELECT A.ID_TIPO_SUMINISTRO,A.PRODUCTO,A.PROCEDIMIENTO,B.NOMBRE,A.ANULADO " & _
                   "  FROM SUMINISTROS_TIPOS A, RPR_TIPOS B " & _
                   " WHERE A.ID_REACTIVO_PR = B.ID_TIPO_REACTIVO_PR " & _
                   filtro & _
                   " ORDER BY A.PRODUCTO"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_ID(ID As Long) As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT A.ID_TIPO_SUMINISTRO,A.PRODUCTO,A.PROCEDIMIENTO,B.NOMBRE,A.ANULADO " & _
                   "  FROM SUMINISTROS_TIPOS A, RPR_TIPOS B " & _
                   " WHERE A.ID_REACTIVO_PR = B.ID_TIPO_REACTIVO_PR " & _
                   "   AND ID_TIPO_SUMINISTRO = " & ID & _
                   " ORDER BY A.PRODUCTO"
        Set Listado_ID = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_TIPO_SUMINISTRO,CONCAT(A.PRODUCTO, ' (',IFNULL(B.CANTIDAD,''),')') " & _
                   "  FROM SUMINISTROS_TIPOS A LEFT JOIN RPR_TIPOS B ON A.ID_REACTIVO_PR = B.ID_TIPO_REACTIVO_PR " & _
                   " ORDER BY A.PRODUCTO"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "SUMINISTROS_TIPOS"
        .setDESCRIPCION = "Tipos de Suministros"
        .setPK = "ID_TIPO_SUMINISTRO"
        .setCAMPO = "PRODUCTO"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
