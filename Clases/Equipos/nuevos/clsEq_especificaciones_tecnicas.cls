VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEq_especificaciones_tecnicas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQ_ESPECIFICACIONES_TECNICAS
'***************************************************************
Private EQUIPO_ID As Long
Private PARAMETRO_ID As Long
Private ORDEN As Long
Private ESPECIFICACION As String
Private REQUISITOS As String
Private OBSERVACIONES As String
Private CONFORME As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQ_ESPECIFICACIONES_TECNICAS
'***************************************************************
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setPARAMETRO_ID(ByVal dato As Long)
        PARAMETRO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setESPECIFICACION(ByVal dato As String)
        ESPECIFICACION = dato
End Property
Public Property Let setREQUISITOS(ByVal dato As String)
        REQUISITOS = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setCONFORME(ByVal dato As Long)
        CONFORME = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQ_ESPECIFICACIONES_TECNICAS
'***************************************************************
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getPARAMETRO_ID() As Long
        getPARAMETRO_ID = PARAMETRO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getESPECIFICACION() As String
        getESPECIFICACION = ESPECIFICACION
End Property
Public Property Get getREQUISITOS() As String
        getREQUISITOS = REQUISITOS
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getCONFORME() As Long
        getCONFORME = CONFORME
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQ_ESPECIFICACIONES_TECNICAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM EQ_ESPECIFICACIONES_TECNICAS WHERE EQUIPO_ID = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                EQUIPO_ID = RS("EQUIPO_ID")
                PARAMETRO_ID = RS("PARAMETRO_ID")
                ORDEN = RS("ORDEN")
                ESPECIFICACION = RS("ESPECIFICACION")
                REQUISITOS = RS("REQUISITOS")
                OBSERVACIONES = RS("OBSERVACIONES")
                CONFORME = RS("CONFORME")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEq_especificaciones_tecnicas)", vbCritical, Err.Description
End Function
Public Function CrearORDEN(EQUIPO As Long)
        Dim RS As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MAX(ORDEN) FROM EQ_ESPECIFICACIONES_TECNICAS WHERE EQUIPO_ID = " & EQUIPO
        Set RS = datos_bd(consulta)
        If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then
                ORDEN = 1
        Else
                ORDEN = RS.Fields(0) + 1
        End If
        Set RS = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearORDEN EQUIPO_ID
        consulta = "INSERT INTO EQ_ESPECIFICACIONES_TECNICAS " & _
                   "  VALUES (" & _
                        EQUIPO_ID & "," & PARAMETRO_ID & "," & ORDEN & "," & _
                        "'" & ESPECIFICACION & "'" & "," & "'" & REQUISITOS & "'" & "," & "'" & OBSERVACIONES & "'" & "," & _
                        CONFORME & _
                ")"
        execute_bd consulta
        Insertar = EQUIPO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEq_especificaciones_tecnicas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQ_ESPECIFICACIONES_TECNICAS SET " & _
                        " PARAMETRO_ID = " & PARAMETRO_ID & "," & _
                        " ORDEN = " & ORDEN & "," & _
                        " ESPECIFICACION = '" & ESPECIFICACION & "'," & _
                        " REQUISITOS = '" & REQUISITOS & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " CONFORME = " & CONFORME & "" & _
                " WHERE EQUIPO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEq_especificaciones_tecnicas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EQ_ESPECIFICACIONES_TECNICAS " & _
                "    WHERE EQUIPO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEq_especificaciones_tecnicas)", vbCritical, Err.Description
End Function
Public Function Listado(EQUIPO As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT B.NOMBRE,C.CODIGO,A.ESPECIFICACION,A.REQUISITOS,A.OBSERVACIONES,A.CONFORME,A.PARAMETRO_ID,B.NORMA_ID " & _
                   "  FROM EQ_ESPECIFICACIONES_TECNICAS A LEFT JOIN EQ_PARAMETROS_TECNICOS B " & _
                   "    ON A.PARAMETRO_ID = B.ID_PARAMETRO " & _
                   " LEFT JOIN CA_NORMAS C ON B.NORMA_ID = C.ID_NORMA " & _
                   " WHERE A.EQUIPO_ID = " & EQUIPO & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
