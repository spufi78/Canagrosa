VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsALODINE_PARAMETROS
'***************************************************************
Private ID_PARAMETRO As Long
Private ALODINE_ID As Long
Private PARAMETRO As String
Private rango As String
Private PROCEDIMIENTO As String
Private UNIDAD_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsALODINE_PARAMETROS
'***************************************************************
Public Property Let setID_PARAMETRO(ByVal dato As Long)
        ID_PARAMETRO = dato
End Property
Public Property Let setALODINE_ID(ByVal dato As Long)
        ALODINE_ID = dato
End Property
Public Property Let setPARAMETRO(ByVal dato As String)
        PARAMETRO = dato
End Property
Public Property Let setRANGO(ByVal dato As String)
        rango = dato
End Property
Public Property Let setPROCEDIMIENTO(ByVal dato As String)
        PROCEDIMIENTO = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Long)
        UNIDAD_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsALODINE_PARAMETROS
'***************************************************************
Public Property Get getID_PARAMETRO() As Long
        getID_PARAMETRO = ID_PARAMETRO
End Property
Public Property Get getALODINE_ID() As Long
        getALODINE_ID = ALODINE_ID
End Property
Public Property Get getPARAMETRO() As String
        getPARAMETRO = PARAMETRO
End Property
Public Property Get getRANGO() As String
        getRANGO = rango
End Property
Public Property Get getPROCEDIMIENTO() As String
        getPROCEDIMIENTO = PROCEDIMIENTO
End Property
Public Property Get getUNIDAD_ID() As Long
        getUNIDAD_ID = UNIDAD_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsALODINE_PARAMETROS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_PARAMETROS WHERE ID_PARAMETRO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PARAMETRO = rs("ID_PARAMETRO")
                ALODINE_ID = rs("ALODINE_ID")
                PARAMETRO = rs("PARAMETRO")
                rango = rs("RANGO")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                UNIDAD_ID = rs("UNIDAD_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsALODINE_PARAMETROS)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_PARAMETRO) FROM ALODINE_PARAMETROS "
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PARAMETRO = 1
        Else
                ID_PARAMETRO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_PARAMETRO = 0 Then
            Me.CrearID
        End If
        consulta = "INSERT INTO ALODINE_PARAMETROS " & _
                   "  VALUES (" & _
                        ID_PARAMETRO & "," & ALODINE_ID & "," & "'" & PARAMETRO & "'" & "," & _
                        "'" & rango & "'" & "," & "'" & PROCEDIMIENTO & "'" & "," & UNIDAD_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_PARAMETRO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsALODINE_PARAMETROS) : " & Err.Description, vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_PARAMETROS SET " & _
                        " ALODINE_ID = " & ALODINE_ID & "," & _
                        " PARAMETRO = '" & PARAMETRO & "'," & _
                        " RANGO = '" & rango & "'," & _
                        " PROCEDIMIENTO = '" & PROCEDIMIENTO & "'," & _
                        " UNIDAD_ID = " & UNIDAD_ID & "" & _
                " WHERE ID_PARAMETRO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsALODINE_PARAMETROS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_PARAMETROS " & _
                "    WHERE ALODINE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsALODINE_PARAMETROS)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_PARAMETROS WHERE ALODINE_ID = " & ID & " ORDER BY ID_PARAMETRO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Parametros(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT A.PARAMETRO,A.RANGO,A.PROCEDIMIENTO,U.NOMBRE,A.UNIDAD_ID, A.ID_PARAMETRO " & _
                   "  FROM ALODINE_PARAMETROS A, UNIDADES U " & _
                   " WHERE A.ALODINE_ID = " & ID & _
                   "   AND A.UNIDAD_ID = U.ID_UNIDAD " & _
                   " ORDER BY A.ID_PARAMETRO"
        Set Listado_Parametros = datos_bd(consulta)
End Function
