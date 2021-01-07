VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_categorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEMPLEADOS_CATEGORIAS
'***************************************************************
Private id_categoria As Long
Private DEPARTAMENTO_ID As Long
Private CODIGO As String
Private DESCRIPCION As String
Private DOCUMENTO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEMPLEADOS_CATEGORIAS
'***************************************************************
Public Property Let setID_CATEGORIA(ByVal dato As Long)
        id_categoria = dato
End Property
Public Property Let setDEPARTAMENTO_ID(ByVal dato As Long)
        DEPARTAMENTO_ID = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEMPLEADOS_CATEGORIAS
'***************************************************************
Public Property Get getID_CATEGORIA() As Long
        getID_CATEGORIA = id_categoria
End Property
Public Property Get getDEPARTAMENTO_ID() As Long
        getDEPARTAMENTO_ID = DEPARTAMENTO_ID
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEMPLEADOS_CATEGORIAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EMPLEADOS_CATEGORIAS WHERE ID_CATEGORIA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                id_categoria = rs("ID_CATEGORIA")
                DEPARTAMENTO_ID = rs("DEPARTAMENTO_ID")
                CODIGO = rs("CODIGO")
                DESCRIPCION = rs("DESCRIPCION")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpleados_categorias)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_CATEGORIA) FROM EMPLEADOS_CATEGORIAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                id_categoria = 1
        Else
                id_categoria = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO EMPLEADOS_CATEGORIAS " & _
                   "  VALUES (" & _
                        id_categoria & "," & DEPARTAMENTO_ID & "," & "'" & CODIGO & "'" & "," & _
                        "'" & DESCRIPCION & "'," & DOCUMENTO_ID & _
                ")"
        execute_bd consulta
        Insertar = id_categoria
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEmpleados_categorias)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EMPLEADOS_CATEGORIAS SET " & _
                        " DEPARTAMENTO_ID = " & DEPARTAMENTO_ID & "," & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE ID_CATEGORIA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEmpleados_categorias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EMPLEADOS_CATEGORIAS " & _
                "    WHERE ID_CATEGORIA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEmpleados_categorias)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.VALOR,A.DESCRIPCION,B.ID_CATEGORIA,CONCAT('(',B.CODIGO,') ',B.DESCRIPCION), B.DOCUMENTO_ID " & _
                    " FROM decodificadora A LEFT JOIN EMPLEADOS_CATEGORIAS B " & _
                    "   ON A.VALOR = B.DEPARTAMENTO_ID " & _
                    " Where a.CODIGO = " & DECODIFICADORA.EMPLEADOS_DEPARTAMENTOS
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_CATEGORIA,DEPARTAMENTO_ID FROM EMPLEADOS_CATEGORIAS ORDER BY DEPARTAMENTO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    Dim c As String
    c = "select a.ID_CATEGORIA,concat(a.CODIGO,'-',a.DESCRIPCION,' (',b.DESCRIPCION,')') " & _
        " from empleados_categorias a, decodificadora b " & _
        " Where a.DEPARTAMENTO_ID = b.valor " & _
        "  and b.CODIGO = " & DECODIFICADORA.EMPLEADOS_DEPARTAMENTOS
    With combo
        .setCONN = conn
        .setQUERY = c
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "EMPLEADOS_CATEGORIAS"
        .setDESCRIPCION = "Departamentos y categorias"
        .setPK = "ID_CATEGORIA"
        .setCAMPO = "concat(a.CODIGO,'-',a.DESCRIPCION,' (',b.DESCRIPCION,')')"
        .setFILTRO = filtro
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function

