VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_ficha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_FICHA
'***************************************************************
Private ID_FICHA As Long
Private RECUBRIMIENTO_ID As Long
Private FABRICANTE_ID As Long
Private OMAT As String
Private metco As String
Private MICROESTRUCTURA As Long
Private TRACCION As Long
Private MACRO_DUREZA As Long
Private MICRO_DUREZA As Long
Private TRACCION_REQ As String
Private MACRO_DUREZA_REQ As String
Private MICRO_DUREZA_REQ As String
Private ESPESOR As Long
Private ESPESOR_REQ As String
Private SHOREA As Long
Private SHOREA_REQ As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_FICHA
'***************************************************************
Public Property Let setID_FICHA(ByVal dato As Long)
        ID_FICHA = dato
End Property
Public Property Let setRECUBRIMIENTO_ID(ByVal dato As Long)
        RECUBRIMIENTO_ID = dato
End Property
Public Property Let setFABRICANTE_ID(ByVal dato As Long)
        FABRICANTE_ID = dato
End Property
Public Property Let setOMAT(ByVal dato As String)
        OMAT = dato
End Property
Public Property Let setMETCO(ByVal dato As String)
        metco = dato
End Property
Public Property Let setMICROESTRUCTURA(ByVal dato As Long)
        MICROESTRUCTURA = dato
End Property
Public Property Let setTRACCION(ByVal dato As Long)
        TRACCION = dato
End Property
Public Property Let setMACRO_DUREZA(ByVal dato As Long)
        MACRO_DUREZA = dato
End Property
Public Property Let setMICRO_DUREZA(ByVal dato As Long)
        MICRO_DUREZA = dato
End Property
Public Property Let setTRACCION_REQ(ByVal dato As String)
        TRACCION_REQ = dato
End Property
Public Property Let setMACRO_DUREZA_REQ(ByVal dato As String)
        MACRO_DUREZA_REQ = dato
End Property
Public Property Let setMICRO_DUREZA_REQ(ByVal dato As String)
        MICRO_DUREZA_REQ = dato
End Property
Public Property Let setESPESOR(ByVal dato As Long)
        ESPESOR = dato
End Property
Public Property Let setESPESOR_REQ(ByVal dato As String)
        ESPESOR_REQ = dato
End Property
Public Property Let setSHOREA(ByVal dato As Long)
        SHOREA = dato
End Property
Public Property Let setSHOREA_REQ(ByVal dato As String)
        SHOREA_REQ = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_FICHA
'***************************************************************
Public Property Get getID_FICHA() As Long
        getID_FICHA = ID_FICHA
End Property
Public Property Get getRECUBRIMIENTO_ID() As Long
        getRECUBRIMIENTO_ID = RECUBRIMIENTO_ID
End Property
Public Property Get getFABRICANTE_ID() As Long
        getFABRICANTE_ID = FABRICANTE_ID
End Property
Public Property Get getOMAT() As String
        getOMAT = OMAT
End Property
Public Property Get getMETCO() As String
        getMETCO = metco
End Property
Public Property Get getMICROESTRUCTURA() As Long
        getMICROESTRUCTURA = MICROESTRUCTURA
End Property
Public Property Get getTRACCION() As Long
        getTRACCION = TRACCION
End Property
Public Property Get getMACRO_DUREZA() As Long
        getMACRO_DUREZA = MACRO_DUREZA
End Property
Public Property Get getMICRO_DUREZA() As Long
        getMICRO_DUREZA = MICRO_DUREZA
End Property
Public Property Get getTRACCION_REQ() As String
        getTRACCION_REQ = TRACCION_REQ
End Property
Public Property Get getMACRO_DUREZA_REQ() As String
        getMACRO_DUREZA_REQ = MACRO_DUREZA_REQ
End Property
Public Property Get getMICRO_DUREZA_REQ() As String
        getMICRO_DUREZA_REQ = MICRO_DUREZA_REQ
End Property
Public Property Get getESPESOR() As Long
        getESPESOR = ESPESOR
End Property
Public Property Get getESPESOR_REQ() As String
        getESPESOR_REQ = ESPESOR_REQ
End Property
Public Property Get getSHOREA() As Long
        getSHOREA = SHOREA
End Property
Public Property Get getSHOREA_REQ() As String
        getSHOREA_REQ = SHOREA_REQ
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_FICHA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_FICHA WHERE ID_FICHA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FICHA = rs("ID_FICHA")
                RECUBRIMIENTO_ID = rs("RECUBRIMIENTO_ID")
                FABRICANTE_ID = rs("FABRICANTE_ID")
                OMAT = rs("OMAT")
                metco = rs("METCO")
                MICROESTRUCTURA = rs("MICROESTRUCTURA")
                TRACCION = rs("TRACCION")
                MACRO_DUREZA = rs("MACRO_DUREZA")
                MICRO_DUREZA = rs("MICRO_DUREZA")
                ESPESOR = rs("ESPESOR")
                TRACCION_REQ = rs("TRACCION_REQ")
                MACRO_DUREZA_REQ = rs("MACRO_DUREZA_REQ")
                MICRO_DUREZA_REQ = rs("MICRO_DUREZA_REQ")
                ESPESOR_REQ = rs("ESPESOR_REQ")
                SHOREA = rs("SHOREA")
                SHOREA_REQ = rs("SHOREA_REQ")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_ficha)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_FICHA) FROM PLASMA_FICHA"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FICHA = 1
        Else
                ID_FICHA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PLASMA_FICHA " & _
                   "  VALUES (" & _
                        ID_FICHA & "," & RECUBRIMIENTO_ID & "," & FABRICANTE_ID & ",'" & OMAT & "'" & "," & _
                        "'" & metco & "'" & "," & MICROESTRUCTURA & "," & _
                        TRACCION & ",'" & TRACCION_REQ & "'," & _
                        MACRO_DUREZA & ",'" & MACRO_DUREZA_REQ & "'," & _
                        MICRO_DUREZA & ",'" & MICRO_DUREZA_REQ & "'," & _
                        ESPESOR & ",'" & ESPESOR_REQ & "'," & _
                        SHOREA & ",'" & SHOREA_REQ & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_FICHA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_ficha)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_FICHA SET " & _
                        " RECUBRIMIENTO_ID = " & RECUBRIMIENTO_ID & "," & _
                        " FABRICANTE_ID = " & FABRICANTE_ID & "," & _
                        " OMAT = '" & OMAT & "'," & _
                        " METCO = '" & metco & "'," & _
                        " MICROESTRUCTURA = " & MICROESTRUCTURA & "," & _
                        " TRACCION = " & TRACCION & "," & _
                        " MACRO_DUREZA = " & MACRO_DUREZA & "," & _
                        " MICRO_DUREZA = " & MICRO_DUREZA & "," & _
                        " ESPESOR = " & ESPESOR & "," & _
                        " SHOREA = " & SHOREA & "," & _
                        " TRACCION_REQ = '" & TRACCION_REQ & "'," & _
                        " MACRO_DUREZA_REQ = '" & MACRO_DUREZA_REQ & "'," & _
                        " MICRO_DUREZA_REQ = '" & MICRO_DUREZA_REQ & "'," & _
                        " ESPESOR_REQ = '" & ESPESOR_REQ & "'," & _
                        " SHOREA_REQ = '" & SHOREA_REQ & "'" & _
                " WHERE ID_FICHA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPlasma_ficha)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PLASMA_FICHA " & _
                "    WHERE ID_FICHA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_ficha)", vbCritical, Err.Description
End Function
Public Function Listado(metco As String, n_fab As String, fabricante As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE 1 = 1 "
        If metco <> "" Then
            filtro = filtro & " and A.METCO LIKE '%" & metco & "%'"
        End If
        If n_fab <> "" Then
            filtro = filtro & " and A.OMAT LIKE '%" & n_fab & "%'"
        End If
        If fabricante <> 0 Then
            filtro = filtro & " and A.FABRICANTE_ID = " & fabricante
        End If
        consulta = "SELECT A.ID_FICHA,B.DESCRIPCION,A.METCO,A.OMAT,C.NOMBRE,D.NOMBRE,E.NOMBRE,F.NOMBRE,G.NOMBRE,H.NOMBRE " & _
                   "  FROM PLASMA_FICHA A " & _
                   "  LEFT JOIN decodificadora B ON A.RECUBRIMIENTO_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PLASMA_RECUBRIMIENTOS & _
                   "  LEFT JOIN PLASMA_ENSAYOS C ON A.MICROESTRUCTURA = C.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS D ON A.TRACCION = D.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS E ON A.MACRO_DUREZA = E.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS F ON A.MICRO_DUREZA = F.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS G ON A.ESPESOR = G.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS H ON A.SHOREA = H.ID_ENSAYO " & _
                   filtro & _
                   " ORDER BY ID_FICHA"
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoId(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_FICHA,B.DESCRIPCION,A.METCO,A.OMAT,C.NOMBRE,D.NOMBRE,E.NOMBRE,F.NOMBRE,G.NOMBRE,H.NOMBRE " & _
                   "  FROM PLASMA_FICHA A " & _
                   "  LEFT JOIN decodificadora B ON A.RECUBRIMIENTO_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PLASMA_RECUBRIMIENTOS & _
                   "  LEFT JOIN PLASMA_ENSAYOS C ON A.MICROESTRUCTURA = C.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS D ON A.TRACCION = D.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS E ON A.MACRO_DUREZA = E.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS F ON A.MICRO_DUREZA = F.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS G ON A.ESPESOR = G.ID_ENSAYO " & _
                   "  LEFT JOIN PLASMA_ENSAYOS H ON A.SHOREA = H.ID_ENSAYO " & _
                   " WHERE A.ID_FICHA = " & ID & _
                   " ORDER BY ID_FICHA"
        Set ListadoId = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_FICHA
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_FICHA,RECUBRIMIENTO_ID FROM PLASMA_FICHA ORDER BY RECUBRIMIENTO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setFILTRO = filtro
        .setTABLA = "PLASMA_FICHA"
        .setDESCRIPCION = "Fichas de Plasma"
        .setPK = "ID_FICHA"
        .setCAMPO = "CONCAT(OMAT, '-', METCO)"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
