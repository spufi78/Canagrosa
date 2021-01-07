VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_dureza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_DUREZA
'***************************************************************
Private MUESTRA_ID As Long
Private IDENTIFICATION As String
Private DIMENSION As String
Private ESPESOR As String
Private REQUIREMENT As String
Private RESULT As String
Private AVERAGE As String
Private SD As String
Private por As String
Private PASS As Long


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_DUREZA
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setIDENTIFICATION(ByVal dato As String)
        IDENTIFICATION = dato
End Property
Public Property Let setDIMENSION(ByVal dato As String)
        DIMENSION = dato
End Property
Public Property Let setESPESOR(ByVal dato As String)
        ESPESOR = dato
End Property
Public Property Let setREQUIREMENT(ByVal dato As String)
        REQUIREMENT = dato
End Property
Public Property Let setRESULT(ByVal dato As String)
        RESULT = dato
End Property
Public Property Let setAVERAGE(ByVal dato As String)
        AVERAGE = dato
End Property
Public Property Let setSD(ByVal dato As String)
        SD = dato
End Property
Public Property Let setPOR(ByVal dato As String)
        por = dato
End Property
Public Property Let setPASS(ByVal dato As Long)
        PASS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_DUREZA
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getIDENTIFICATION() As String
        getIDENTIFICATION = IDENTIFICATION
End Property
Public Property Get getDIMENSION() As String
        getDIMENSION = DIMENSION
End Property
Public Property Get getESPESOR() As String
        getESPESOR = ESPESOR
End Property
Public Property Get getREQUIREMENT() As String
        getREQUIREMENT = REQUIREMENT
End Property
Public Property Get getRESULT() As String
        getRESULT = RESULT
End Property
Public Property Get getAVERAGE() As String
        getAVERAGE = AVERAGE
End Property
Public Property Get getSD() As String
        getSD = SD
End Property
Public Property Get getPOR() As String
        getPOR = por
End Property
Public Property Get getPASS() As Long
        getPASS = PASS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_DUREZA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_DUREZA WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                IDENTIFICATION = rs("IDENTIFICATION")
                DIMENSION = rs("DIMENSION")
                ESPESOR = rs("ESPESOR")
                REQUIREMENT = rs("REQUIREMENT")
                RESULT = rs("RESULT")
                AVERAGE = rs("AVERAGE")
                SD = rs("SD")
                por = rs("POR")
                PASS = rs("PASS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_dureza)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE INTO PLASMA_DUREZA " & _
                   "  VALUES (" & _
                        MUESTRA_ID & ",'" & IDENTIFICATION & "','" & DIMENSION & "','" & ESPESOR & "','" & REQUIREMENT & "'" & "," & _
                        "'" & RESULT & "'" & "," & "'" & AVERAGE & "','" & SD & "','" & por & "'," & PASS & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_dureza)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_DUREZA SET " & _
                        " IDENTIFICATION = '" & IDENTIFICATION & "'," & _
                        " DIMENSION = '" & DIMENSION & "'," & _
                        " ESPESOR = '" & ESPESOR & "'," & _
                        " REQUIREMENT = '" & REQUIREMENT & "'," & _
                        " RESULT = '" & RESULT & "'," & _
                        " AVERAGE = '" & AVERAGE & "'," & _
                        " SD = '" & SD & "'," & _
                        " POR = '" & por & "'," & _
                        " PASS = " & PASS & "" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPlasma_dureza)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PLASMA_DUREZA " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_dureza)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_DUREZA ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_DUREZA
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

