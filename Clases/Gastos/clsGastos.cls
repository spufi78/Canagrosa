VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSGASTOS
'***************************************************************
Private ID_GASTO As Long
Private PROVEEDOR_ID As Long
Private BANCO_ID As Long
Private TIPO_ID As Long
Private fecha As String
Private DESCRIPCION As String
Private IMPORTE As String
Private FP_ID As Long
Private OBSERVACIONES As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSGASTOS
'***************************************************************
Public Property Let setID_GASTO(ByVal dato As Long)
        ID_GASTO = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setBANCO_ID(ByVal dato As Long)
        BANCO_ID = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
Public Property Let setFP_ID(ByVal dato As Long)
        FP_ID = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSGASTOS
'***************************************************************
Public Property Get getID_GASTO() As Long
        getID_GASTO = ID_GASTO
End Property
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getBANCO_ID() As Long
        getBANCO_ID = BANCO_ID
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
Public Property Get getFP_ID() As Long
        getFP_ID = FP_ID
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSGASTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM GASTOS WHERE ID_GASTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_GASTO = rs("ID_GASTO")
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                BANCO_ID = rs("BANCO_ID")
                TIPO_ID = rs("TIPO_ID")
                fecha = rs("FECHA")
                DESCRIPCION = rs("DESCRIPCION")
                IMPORTE = rs("IMPORTE")
                FP_ID = rs("FP_ID")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsGastos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_GASTO) FROM GASTOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_GASTO = 1
        Else
                ID_GASTO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO GASTOS " & _
                   "  VALUES (" & _
                        ID_GASTO & "," & PROVEEDOR_ID & "," & BANCO_ID & "," & _
                        TIPO_ID & "," & "'" & fecha & "'" & "," & "'" & DESCRIPCION & "'" & ",'" & _
                        IMPORTE & "'," & FP_ID & "," & "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_GASTO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsGastos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE GASTOS SET " & _
                        " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                        " BANCO_ID = " & BANCO_ID & "," & _
                        " TIPO_ID = " & TIPO_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " IMPORTE = '" & IMPORTE & "'," & _
                        " FP_ID = " & FP_ID & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE ID_GASTO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsGastos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM GASTOS " & _
                "    WHERE ID_GASTO = " & ID
        execute_bd consulta
        Dim oD As New clsDocumentacion
        oD.EliminarGasto ID
        Set oD = Nothing
        
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsGastos)", vbCritical, Err.Description
End Function
Public Function Listado(proveedor As Long, banco As Long, tipo As Long, DESCRIPCION As String, fdesde As String, fhasta As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE 1 = 1"
        If proveedor <> 0 Then
            filtro = filtro & " AND A.PROVEEDOR_ID = " & proveedor
        End If
        If banco <> 0 Then
            filtro = filtro & " AND A.BANCO_ID = " & banco
        End If
        If tipo <> 0 Then
            filtro = filtro & " AND A.TIPO_ID = " & tipo
        End If
        If DESCRIPCION <> "" Then
            filtro = filtro & " AND A.DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
        End If
        filtro = filtro & " AND A.FECHA >='" & Format(fdesde, "yyyy-mm-dd") & "'"
        filtro = filtro & " AND A.FECHA <='" & Format(fhasta, "yyyy-mm-dd") & "'"
        
        consulta = "SELECT A.ID_GASTO,D.DESCRIPCION,A.FECHA,A.DESCRIPCION,B.NOMBRE,C.DESCRIPCION,E.NOMBRE,A.IMPORTE " & _
                   "  FROM GASTOS A " & _
                   "  LEFT JOIN PROVEEDORES B ON A.PROVEEDOR_ID = B.ID_PROVEEDOR " & _
                   "  LEFT JOIN BANCOS C ON A.BANCO_ID = C.ID_BANCO " & _
                   "  LEFT JOIN decodificadora D ON A.TIPO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_GASTOS_TIPOS & _
                   "  LEFT JOIN FORMAS_PAGO E ON A.FP_ID = E.ID_FP " & _
                   filtro & _
                   " ORDER BY A.ID_GASTO "
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoId(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_GASTO,D.DESCRIPCION,A.FECHA,A.DESCRIPCION,B.NOMBRE,C.DESCRIPCION,E.NOMBRE,A.IMPORTE " & _
                   "  FROM GASTOS A " & _
                   "  LEFT JOIN PROVEEDORES B ON A.PROVEEDOR_ID = B.ID_PROVEEDOR " & _
                   "  LEFT JOIN BANCOS C ON A.BANCO_ID = C.ID_BANCO " & _
                   "  LEFT JOIN decodificadora D ON A.TIPO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_GASTOS_TIPOS & _
                   "  LEFT JOIN FORMAS_PAGO E ON A.FP_ID = E.ID_FP " & _
                   " WHERE ID_GASTO = " & ID & _
                   " ORDER BY A.ID_GASTO "
        Set ListadoId = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSGASTOS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_GASTO,PROVEEDOR_ID FROM GASTOS ORDER BY PROVEEDOR_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
