VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_ficha_estructura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_FICHA_ESTRUCTURA
'***************************************************************
Private FICHA_ID As Long
Private ENSAYO_ID As Long
Private REQUIREMENT As String


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_FICHA_ESTRUCTURA
'***************************************************************
Public Property Let setFICHA_ID(ByVal dato As Long)
        FICHA_ID = dato
End Property
Public Property Let setENSAYO_ID(ByVal dato As Long)
        ENSAYO_ID = dato
End Property
Public Property Let setREQUIREMENT(ByVal dato As String)
        REQUIREMENT = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_FICHA_ESTRUCTURA
'***************************************************************
Public Property Get getFICHA_ID() As Long
        getFICHA_ID = FICHA_ID
End Property
Public Property Get getENSAYO_ID() As Long
        getENSAYO_ID = ENSAYO_ID
End Property
Public Property Get getREQUIREMENT() As String
        getREQUIREMENT = REQUIREMENT
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_FICHA_ESTRUCTURA
'***************************************************************
Public Function Carga(ID As Long, ENSAYO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_FICHA_ESTRUCTURA WHERE FICHA_ID = " & ID & " AND ENSAYO_ID = " & ENSAYO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                FICHA_ID = rs("FICHA_ID")
                ENSAYO_ID = rs("ENSAYO_ID")
                REQUIREMENT = rs("REQUIREMENT")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_ficha_estructura)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PLASMA_FICHA_ESTRUCTURA " & _
                   "  VALUES (" & _
                        FICHA_ID & "," & ENSAYO_ID & "," & "'" & REQUIREMENT & "'" & _
                ")"
        execute_bd consulta
        Insertar = FICHA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_ficha_estructura)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_FICHA_ESTRUCTURA SET " & _
                        " ENSAYO_ID = " & ENSAYO_ID & "," & _
                        " REQUIREMENT = '" & REQUIREMENT & "'" & _
                " WHERE FICHA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPlasma_ficha_estructura)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PLASMA_FICHA_ESTRUCTURA " & _
                "    WHERE FICHA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_ficha_estructura)", vbCritical, Err.Description
End Function
Public Function Listado(FICHA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.VALOR, A.DESCRIPCION, B.REQUIREMENT " & _
                   "  FROM decodificadora A " & _
                   "  LEFT JOIN PLASMA_FICHA_ESTRUCTURA B ON B.FICHA_ID = " & FICHA & " AND B.ENSAYO_ID = A.VALOR " & _
                   " WHERE A.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PLASMA_DETERMINACIONES & _
                   " ORDER BY A.VALOR"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_FICHA_ESTRUCTURA
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function ListadoCompleto(FICHA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM PLASMA_FICHA_ESTRUCTURA " & _
                   " WHERE FICHA_ID = " & FICHA & _
                   " ORDER BY ENSAYO_ID "
        Set ListadoCompleto = datos_bd(consulta)
End Function


