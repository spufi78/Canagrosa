VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSEGUROS
'***************************************************************
Private ID_SEGURO As Long
Private DESCRIPCION As String
Private poliza As String
Private PROVEEDOR_ID As Long
Private BANCO_ID As Long
Private F_ALTA As String
Private F_VENCIMIENTO As String
Private PERIODICIDAD_ID As Long
Private IMPORTE As String
Private SUBCUENTA As String

Private OBSERVACIONES As String


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSEGUROS
'***************************************************************
Public Property Let setID_SEGURO(ByVal dato As Long)
        ID_SEGURO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setPOLIZA(ByVal dato As String)
        poliza = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setBANCO_ID(ByVal dato As Long)
        BANCO_ID = dato
End Property
Public Property Let setF_ALTA(ByVal dato As String)
        F_ALTA = dato
End Property
Public Property Let setF_VENCIMIENTO(ByVal dato As String)
        F_VENCIMIENTO = dato
End Property
Public Property Let setPERIODICIDAD_ID(ByVal dato As Long)
        PERIODICIDAD_ID = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
Public Property Let setSUBCUENTA(ByVal dato As String)
        SUBCUENTA = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSEGUROS
'***************************************************************
Public Property Get getID_SEGURO() As Long
        getID_SEGURO = ID_SEGURO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getPOLIZA() As String
        getPOLIZA = poliza
End Property
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getBANCO_ID() As Long
        getBANCO_ID = BANCO_ID
End Property
Public Property Get getF_ALTA() As String
        getF_ALTA = F_ALTA
End Property
Public Property Get getF_VENCIMIENTO() As String
        getF_VENCIMIENTO = F_VENCIMIENTO
End Property
Public Property Get getPERIODICIDAD_ID() As Long
        getPERIODICIDAD_ID = PERIODICIDAD_ID
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
Public Property Get getSUBCUENTA() As String
        getSUBCUENTA = SUBCUENTA
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSEGUROS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SEGUROS WHERE ID_SEGURO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SEGURO = rs("ID_SEGURO")
                DESCRIPCION = rs("DESCRIPCION")
                poliza = rs("POLIZA")
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                BANCO_ID = rs("BANCO_ID")
                F_ALTA = rs("F_ALTA")
                F_VENCIMIENTO = rs("F_VENCIMIENTO")
                PERIODICIDAD_ID = rs("PERIODICIDAD_ID")
                IMPORTE = rs("IMPORTE")
                SUBCUENTA = rs("SUBCUENTA")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSeguros)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_SEGURO) FROM SEGUROS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_SEGURO = 1
        Else
                ID_SEGURO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO SEGUROS " & _
                   "  VALUES (" & _
                        ID_SEGURO & "," & "'" & DESCRIPCION & "','" & poliza & "'," & PROVEEDOR_ID & "," & BANCO_ID & "," & _
                        "'" & F_ALTA & "'" & "," & "'" & F_VENCIMIENTO & "'" & "," & PERIODICIDAD_ID & ",'" & IMPORTE & "'," & _
                        "'" & SUBCUENTA & "','" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_SEGURO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSeguros)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SEGUROS SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " POLIZA = '" & poliza & "'," & _
                        " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                        " BANCO_ID = " & BANCO_ID & "," & _
                        " F_ALTA = '" & F_ALTA & "'," & _
                        " F_VENCIMIENTO = '" & F_VENCIMIENTO & "'," & _
                        " PERIODICIDAD_ID = " & PERIODICIDAD_ID & "," & _
                        " IMPORTE = '" & IMPORTE & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " SUBCUENTA = '" & SUBCUENTA & "'" & _
                " WHERE ID_SEGURO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSeguros)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim oTP As New clsTesoreria_prevision
        oTP.Eliminar TOBJETO.tobjeto_seguro, ID
        execute_bd "DELETE FROM SEGUROS WHERE ID_SEGURO = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSeguros)", vbCritical, Err.Description
End Function
Public Function Listado(DESCRIPCION As String, poliza As String, proveedor As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE 1 = 1 "
        If DESCRIPCION <> "" Then
            filtro = filtro & " AND A.DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
        End If
        If poliza <> "" Then
            filtro = filtro & " AND A.POLIZA LIKE '%" & poliza & "%'"
        End If
        If proveedor <> 0 Then
            filtro = filtro & " AND A.PROVEEDOR_ID = " & proveedor
        End If
        consulta = "SELECT A.ID_SEGURO,A.POLIZA,A.DESCRIPCION,B.NOMBRE,A.F_ALTA,A.F_VENCIMIENTO,C.DESCRIPCION,A.SUBCUENTA " & _
                   "  FROM SEGUROS A " & _
                   "  LEFT JOIN PROVEEDORES B ON A.PROVEEDOR_ID = B.ID_PROVEEDOR " & _
                   "  LEFT JOIN decodificadora C ON A.PERIODICIDAD_ID = C.VALOR AND C.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PERIODICIDADES_FACTURACION & _
                   filtro & _
                   " ORDER BY ID_SEGURO"
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoId(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_SEGURO,A.POLIZA,A.DESCRIPCION,B.NOMBRE,A.F_ALTA,A.F_VENCIMIENTO,C.DESCRIPCION,A.SUBCUENTA " & _
                   "  FROM SEGUROS A " & _
                   "  LEFT JOIN PROVEEDORES B ON A.PROVEEDOR_ID = B.ID_PROVEEDOR " & _
                   "  LEFT JOIN decodificadora C ON A.PERIODICIDAD_ID = C.VALOR AND C.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PERIODICIDADES_FACTURACION & _
                   " WHERE A.ID_SEGURO = " & ID & _
                   " ORDER BY ID_SEGURO"
        Set ListadoId = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSSEGUROS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_SEGURO,DESCRIPCION FROM SEGUROS ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
