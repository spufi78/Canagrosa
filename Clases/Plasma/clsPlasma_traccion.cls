VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_traccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_TRACCION
'***************************************************************
Private MUESTRA_ID As Long
Private tipo As Long
Private ROOM As String
Private VELOCITY As String
Private EPOXY As String
Private ADHESIVE As String
Private ESPESOR As String
Private AVERAGE As String
Private SD As String


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_TRACCION
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
Public Property Let setROOM(ByVal dato As String)
        ROOM = dato
End Property
Public Property Let setVELOCITY(ByVal dato As String)
        VELOCITY = dato
End Property
Public Property Let setEPOXY(ByVal dato As String)
        EPOXY = dato
End Property
Public Property Let setADHESIVE(ByVal dato As String)
        ADHESIVE = dato
End Property
Public Property Let setESPESOR(ByVal dato As String)
        ESPESOR = dato
End Property
Public Property Let setAVERAGE(ByVal dato As String)
        AVERAGE = dato
End Property
Public Property Let setSD(ByVal dato As String)
        SD = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_TRACCION
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
Public Property Get getROOM() As String
        getROOM = ROOM
End Property
Public Property Get getVELOCITY() As String
        getVELOCITY = VELOCITY
End Property
Public Property Get getEPOXY() As String
        getEPOXY = EPOXY
End Property
Public Property Get getADHESIVE() As String
        getADHESIVE = ADHESIVE
End Property
Public Property Get getESPESOR() As String
        getESPESOR = ESPESOR
End Property
Public Property Get getAVERAGE() As String
        getAVERAGE = AVERAGE
End Property
Public Property Get getSD() As String
        getSD = SD
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_TRACCION
'***************************************************************
Public Function Carga(MUESTRA_ID As Long, tipo As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_TRACCION " & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID & _
                   "   AND TIPO = " & tipo
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                tipo = rs("TIPO")
                ROOM = rs("ROOM")
                VELOCITY = rs("VELOCITY")
                EPOXY = rs("EPOXY")
                ADHESIVE = rs("ADHESIVE")
                ESPESOR = rs("ESPESOR")
                AVERAGE = rs("AVERAGE")
                SD = rs("SD")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_traccion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PLASMA_TRACCION " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & tipo & "," & "'" & ROOM & "','" & VELOCITY & "','" & EPOXY & "','" & ADHESIVE & "','" & _
                        ESPESOR & "','" & AVERAGE & "','" & SD & "'" & _
                ") ON DUPLICATE KEY UPDATE " & _
                        " ROOM = '" & ROOM & "'," & _
                        " EPOXY = '" & EPOXY & "'," & _
                        " ADHESIVE = '" & ADHESIVE & "'," & _
                        " ESPESOR = '" & ESPESOR & "'," & _
                        " VELOCITY = '" & VELOCITY & "'," & _
                        " AVERAGE = '" & AVERAGE & "'," & _
                        " SD = '" & SD & "'"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_traccion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_TRACCION SET " & _
                        " TIPO = " & tipo & "," & _
                        " ROOM = '" & ROOM & "'," & _
                        " VELOCITY = '" & VELOCITY & "'," & _
                        " EPOXY = '" & EPOXY & "'," & _
                        " ADHESIVE = '" & ADHESIVE & "'," & _
                        " ESPESOR = '" & ESPESOR & "'," & _
                        " AVERAGE = " & AVERAGE & "," & _
                        " SD = " & SD & "" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPlasma_traccion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PLASMA_TRACCION " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_traccion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_TRACCION ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_TRACCION
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

