VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_PROCESOS
'***************************************************************
Private ID_PROCESO As Long
Private FABRICANTE_ID As Long
Private nombre As String
Private tipo As Integer
Private BOND_COAT_FICHA_ID As Long
Private BOND_SPECIFICATION As String
Private BOND_MICROESTRUCTURA As Integer
Private BOND_TRACCION As Integer
Private BOND_MACRO_DUREZA As Integer
Private BOND_MICRO_DUREZA As Integer
Private BOND_ESPESOR As Integer
Private TOP_COAT_FICHA_ID As Long
Private TOP_SPECIFICATION As String
Private TOP_MICROESTRUCTURA As Integer
Private TOP_TRACCION As Integer
Private TOP_MACRO_DUREZA As Integer
Private TOP_MICRO_DUREZA As Integer
Private TOP_ESPESOR As Integer


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_PROCESOS
'***************************************************************
Public Property Let setID_PROCESO(ByVal dato As Long)
        ID_PROCESO = dato
End Property
Public Property Let setFABRICANTE_ID(ByVal dato As Long)
        FABRICANTE_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setTIPO(ByVal dato As Integer)
        tipo = dato
End Property
Public Property Let setBOND_COAT_FICHA_ID(ByVal dato As Long)
        BOND_COAT_FICHA_ID = dato
End Property
Public Property Let setBOND_SPECIFICATION(ByVal dato As String)
        BOND_SPECIFICATION = dato
End Property
Public Property Let setBOND_MICROESTRUCTURA(ByVal dato As Integer)
        BOND_MICROESTRUCTURA = dato
End Property
Public Property Let setBOND_TRACCION(ByVal dato As Integer)
        BOND_TRACCION = dato
End Property
Public Property Let setBOND_MACRO_DUREZA(ByVal dato As Integer)
        BOND_MACRO_DUREZA = dato
End Property
Public Property Let setBOND_MICRO_DUREZA(ByVal dato As Integer)
        BOND_MICRO_DUREZA = dato
End Property
Public Property Let setBOND_ESPESOR(ByVal dato As Integer)
        BOND_ESPESOR = dato
End Property
Public Property Let setTOP_COAT_FICHA_ID(ByVal dato As Long)
        TOP_COAT_FICHA_ID = dato
End Property
Public Property Let setTOP_SPECIFICATION(ByVal dato As String)
        TOP_SPECIFICATION = dato
End Property
Public Property Let setTOP_MICROESTRUCTURA(ByVal dato As Integer)
        TOP_MICROESTRUCTURA = dato
End Property
Public Property Let setTOP_TRACCION(ByVal dato As Integer)
        TOP_TRACCION = dato
End Property
Public Property Let setTOP_MACRO_DUREZA(ByVal dato As Integer)
        TOP_MACRO_DUREZA = dato
End Property
Public Property Let setTOP_MICRO_DUREZA(ByVal dato As Integer)
        TOP_MICRO_DUREZA = dato
End Property
Public Property Let setTOP_ESPESOR(ByVal dato As Integer)
        TOP_ESPESOR = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_PROCESOS
'***************************************************************
Public Property Get getID_PROCESO() As Long
        getID_PROCESO = ID_PROCESO
End Property
Public Property Get getFABRICANTE_ID() As Long
        getFABRICANTE_ID = FABRICANTE_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getTIPO() As Integer
        getTIPO = tipo
End Property
Public Property Get getBOND_COAT_FICHA_ID() As Long
        getBOND_COAT_FICHA_ID = BOND_COAT_FICHA_ID
End Property
Public Property Get getBOND_SPECIFICATION() As String
        getBOND_SPECIFICATION = BOND_SPECIFICATION
End Property
Public Property Get getBOND_MICROESTRUCTURA() As Integer
        getBOND_MICROESTRUCTURA = BOND_MICROESTRUCTURA
End Property
Public Property Get getBOND_TRACCION() As Integer
        getBOND_TRACCION = BOND_TRACCION
End Property
Public Property Get getBOND_MACRO_DUREZA() As Integer
        getBOND_MACRO_DUREZA = BOND_MACRO_DUREZA
End Property
Public Property Get getBOND_MICRO_DUREZA() As Integer
        getBOND_MICRO_DUREZA = BOND_MICRO_DUREZA
End Property
Public Property Get getBOND_ESPESOR() As Integer
        getBOND_ESPESOR = BOND_ESPESOR
End Property
Public Property Get getTOP_COAT_FICHA_ID() As Long
        getTOP_COAT_FICHA_ID = TOP_COAT_FICHA_ID
End Property
Public Property Get getTOP_SPECIFICATION() As String
        getTOP_SPECIFICATION = TOP_SPECIFICATION
End Property

Public Property Get getTOP_MICROESTRUCTURA() As Integer
        getTOP_MICROESTRUCTURA = TOP_MICROESTRUCTURA
End Property
Public Property Get getTOP_TRACCION() As Integer
        getTOP_TRACCION = TOP_TRACCION
End Property
Public Property Get getTOP_MACRO_DUREZA() As Integer
        getTOP_MACRO_DUREZA = TOP_MACRO_DUREZA
End Property
Public Property Get getTOP_MICRO_DUREZA() As Integer
        getTOP_MICRO_DUREZA = TOP_MICRO_DUREZA
End Property
Public Property Get getTOP_ESPESOR() As Integer
        getTOP_ESPESOR = TOP_ESPESOR
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_PROCESOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_PROCESOS WHERE ID_PROCESO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PROCESO = rs("ID_PROCESO")
                FABRICANTE_ID = rs("FABRICANTE_ID")
                nombre = rs("NOMBRE")
                tipo = rs("TIPO")
                BOND_COAT_FICHA_ID = rs("BOND_COAT_FICHA_ID")
                BOND_SPECIFICATION = rs("BOND_SPECIFICATION")
                BOND_MICROESTRUCTURA = rs("BOND_MICROESTRUCTURA")
                BOND_TRACCION = rs("BOND_TRACCION")
                BOND_MACRO_DUREZA = rs("BOND_MACRO_DUREZA")
                BOND_MICRO_DUREZA = rs("BOND_MICRO_DUREZA")
                BOND_ESPESOR = rs("BOND_ESPESOR")
                TOP_COAT_FICHA_ID = rs("TOP_COAT_FICHA_ID")
                TOP_SPECIFICATION = rs("TOP_SPECIFICATION")
                TOP_MICROESTRUCTURA = rs("TOP_MICROESTRUCTURA")
                TOP_TRACCION = rs("TOP_TRACCION")
                TOP_MACRO_DUREZA = rs("TOP_MACRO_DUREZA")
                TOP_MICRO_DUREZA = rs("TOP_MICRO_DUREZA")
                TOP_ESPESOR = rs("TOP_ESPESOR")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_procesos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_PROCESO) FROM PLASMA_PROCESOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PROCESO = 1
        Else
                ID_PROCESO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PLASMA_PROCESOS " & _
                   "  VALUES (" & _
                        ID_PROCESO & "," & FABRICANTE_ID & ",'" & nombre & "'," & tipo & "," & BOND_COAT_FICHA_ID & ",'" & BOND_SPECIFICATION & "'," & _
                        BOND_MICROESTRUCTURA & "," & BOND_TRACCION & "," & BOND_MACRO_DUREZA & "," & BOND_MICRO_DUREZA & "," & BOND_ESPESOR & "," & _
                        TOP_COAT_FICHA_ID & ",'" & TOP_SPECIFICATION & "'," & _
                        TOP_MICROESTRUCTURA & "," & TOP_TRACCION & "," & TOP_MACRO_DUREZA & "," & TOP_MICRO_DUREZA & "," & TOP_ESPESOR & _
                ")"
        execute_bd consulta
        Insertar = ID_PROCESO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_procesos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_PROCESOS SET " & _
                        " NOMBRE = '" & nombre & "'," & _
                        " TIPO = " & tipo & "," & _
                        " FABRICANTE_ID = " & FABRICANTE_ID & "," & _
                        " BOND_COAT_FICHA_ID = " & BOND_COAT_FICHA_ID & "," & _
                        " BOND_SPECIFICATION = '" & BOND_SPECIFICATION & "'," & _
                        " BOND_MICROESTRUCTURA = " & BOND_MICROESTRUCTURA & "," & _
                        " BOND_TRACCION = " & BOND_TRACCION & "," & _
                        " BOND_MACRO_DUREZA = " & BOND_MACRO_DUREZA & "," & _
                        " BOND_MICRO_DUREZA = " & BOND_MICRO_DUREZA & "," & _
                        " BOND_ESPESOR = " & BOND_ESPESOR & "," & _
                        " TOP_COAT_FICHA_ID = " & TOP_COAT_FICHA_ID & "," & _
                        " TOP_SPECIFICATION = '" & TOP_SPECIFICATION & "'," & _
                        " TOP_MICROESTRUCTURA = " & TOP_MICROESTRUCTURA & "," & _
                        " TOP_TRACCION = " & TOP_TRACCION & "," & _
                        " TOP_MACRO_DUREZA = " & TOP_MACRO_DUREZA & "," & _
                        " TOP_MICRO_DUREZA = " & TOP_MICRO_DUREZA & "," & _
                        " TOP_ESPESOR = " & TOP_ESPESOR & _
                " WHERE ID_PROCESO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPlasma_procesos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        execute_bd "DELETE FROM PLASMA_PROCESOS WHERE ID_PROCESO = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_procesos)", vbCritical, Err.Description
End Function
Public Function Listado(PROCESO As String, bond As Long, top As Long, tipo As Integer) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE 1 = 1 "
        If tipo <> 0 Then
            filtro = filtro & " AND A.TIPO = " & tipo
        End If
        If PROCESO <> "" Then
            filtro = filtro & " AND A.NOMBRE LIKE '%" & PROCESO & "%'"
        End If
        If bond <> 0 Then
            filtro = filtro & " AND A.BOND_COAT_FICHA_ID = " & bond
        End If
        If top <> 0 Then
            filtro = filtro & " AND A.TOP_COAT_FICHA_ID = " & top
        End If
        consulta = "SELECT A.ID_PROCESO,A.NOMBRE,FAB.DESCRIPCION,CONCAT(IFNULL(B.METCO,''),' ',IFNULL(B1.DESCRIPCION,'')),CONCAT(IFNULL(C.METCO,''),' ',IFNULL(C1.DESCRIPCION,'')) " & _
                   "  FROM PLASMA_PROCESOS A " & _
                   "  LEFT JOIN PLASMA_FICHA B ON A.BOND_COAT_FICHA_ID = B.ID_FICHA " & _
                   "  LEFT JOIN decodificadora B1 ON B.RECUBRIMIENTO_ID = B1.VALOR AND B1.CODIGO=" & DECODIFICADORA.DECODIFICADORA_PLASMA_RECUBRIMIENTOS & _
                   "  LEFT JOIN PLASMA_FICHA C ON A.TOP_COAT_FICHA_ID = C.ID_FICHA " & _
                   "  LEFT JOIN decodificadora C1 ON C.RECUBRIMIENTO_ID = C1.VALOR AND C1.CODIGO=" & DECODIFICADORA.DECODIFICADORA_PLASMA_RECUBRIMIENTOS & _
                   "  LEFT JOIN decodificadora FAB ON A.FABRICANTE_ID = FAB.VALOR AND FAB.CODIGO=" & DECODIFICADORA.DECODIFICADORA_PLASMA_FABRICANTES & _
                   filtro & _
                   " ORDER BY A.NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoId(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_PROCESO,A.NOMBRE,FAB.DESCRIPCION,CONCAT(IFNULL(B.METCO,''),' ',IFNULL(B1.DESCRIPCION,'')),CONCAT(IFNULL(C.METCO,''),' ',IFNULL(C1.DESCRIPCION,'')) " & _
                   "  FROM PLASMA_PROCESOS A " & _
                   "  LEFT JOIN PLASMA_FICHA B ON A.BOND_COAT_FICHA_ID = B.ID_FICHA " & _
                   "  LEFT JOIN decodificadora B1 ON B.RECUBRIMIENTO_ID = B1.VALOR AND B1.CODIGO=" & DECODIFICADORA.DECODIFICADORA_PLASMA_RECUBRIMIENTOS & _
                   "  LEFT JOIN PLASMA_FICHA C ON A.TOP_COAT_FICHA_ID = C.ID_FICHA " & _
                   "  LEFT JOIN decodificadora C1 ON C.RECUBRIMIENTO_ID = C1.VALOR AND C1.CODIGO=" & DECODIFICADORA.DECODIFICADORA_PLASMA_RECUBRIMIENTOS & _
                   "  LEFT JOIN decodificadora FAB ON A.FABRICANTE_ID = FAB.VALOR AND FAB.CODIGO=" & DECODIFICADORA.DECODIFICADORA_PLASMA_FABRICANTES & _
                   " WHERE ID_PROCESO = " & ID & _
                   " ORDER BY A.NOMBRE"
        Set ListadoId = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_PROCESOS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PROCESO,NOMBRE FROM PLASMA_PROCESOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setFILTRO = filtro
        .setTABLA = "PLASMA_PROCESOS"
        .setDESCRIPCION = "Procesos de Plasma"
        .setPK = "ID_PROCESO"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
