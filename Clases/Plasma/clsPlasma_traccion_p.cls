VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_traccion_p"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_TRACCION_P
'***************************************************************
Private MUESTRA_ID As Long
Private tipo As Long
Private ORDEN As Long
Private IDENTIFICATION As String
Private DIAMETER As String
Private area As String
Private LOADP As String
Private TENSILE As String
Private LOCATION As String


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_TRACCION_P
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setIDENTIFICATION(ByVal dato As String)
        IDENTIFICATION = dato
End Property
Public Property Let setDIAMETER(ByVal dato As String)
        DIAMETER = dato
End Property
Public Property Let setAREA(ByVal dato As String)
        area = dato
End Property
Public Property Let setLOADP(ByVal dato As String)
        LOADP = dato
End Property
Public Property Let setTENSILE(ByVal dato As String)
        TENSILE = dato
End Property
Public Property Let setLOCATION(ByVal dato As String)
        LOCATION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_TRACCION_P
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getIDENTIFICATION() As String
        getIDENTIFICATION = IDENTIFICATION
End Property
Public Property Get getDIAMETER() As String
        getDIAMETER = DIAMETER
End Property
Public Property Get getAREA() As String
        getAREA = area
End Property
Public Property Get getLOADP() As String
        getLOADP = LOADP
End Property
Public Property Get getTENSILE() As String
        getTENSILE = TENSILE
End Property
Public Property Get getLOCATION() As String
        getLOCATION = LOCATION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_TRACCION_P
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_TRACCION_P WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                tipo = rs("TIPO")
                ORDEN = rs("ORDEN")
                IDENTIFICATION = rs("IDENTIFICATION")
                DIAMETER = rs("DIAMETER")
                area = rs("AREA")
                LOADP = rs("LOADP")
                TENSILE = rs("TENSILE")
                LOCATION = rs("LOCATION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_traccion_p)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PLASMA_TRACCION_P " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & tipo & "," & ORDEN & "," & _
                        "'" & IDENTIFICATION & "'" & ",'" & DIAMETER & "','" & area & "','" & _
                        LOADP & "','" & TENSILE & "'," & "'" & LOCATION & "'" & _
                   ") ON DUPLICATE KEY UPDATE " & _
                        " IDENTIFICATION = '" & IDENTIFICATION & "'," & _
                        " DIAMETER = '" & DIAMETER & "'," & _
                        " AREA = '" & area & "'," & _
                        " LOADP = '" & LOADP & "'," & _
                        " TENSILE = '" & TENSILE & "'," & _
                        " LOCATION = '" & LOCATION & "'"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_traccion_p)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PLASMA_TRACCION_P " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_traccion_p)", vbCritical, Err.Description
End Function
Public Function Listado(MUESTRA_ID As Long, tipo As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_TRACCION_P " & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID & _
                   "   AND TIPO = " & tipo & _
                   " ORDER BY ORDEN "
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_TRACCION_P
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

