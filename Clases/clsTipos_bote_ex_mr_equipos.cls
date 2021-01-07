VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_bote_ex_mr_equipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_BOTE_EX_MR_EQUIPOS
'***************************************************************
Private TIPO_BOTE_EX_ID As Long
Private EQUIPO_ID As Long
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_BOTE_EX_MR_EQUIPOS
'***************************************************************
Public Property Let setTIPO_BOTE_EX_ID(ByVal dato As Long)
        TIPO_BOTE_EX_ID = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_BOTE_EX_MR_EQUIPOS
'***************************************************************
Public Property Get getTIPO_BOTE_EX_ID() As Long
        getTIPO_BOTE_EX_ID = TIPO_BOTE_EX_ID
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTIPOS_BOTE_EX_MR_EQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_BOTE_EX_MR_EQUIPOS WHERE TIPO_BOTE_EX_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TIPO_BOTE_EX_ID = rs("TIPO_BOTE_EX_ID")
                EQUIPO_ID = rs("EQUIPO_ID")
                FTIMESTP = rs("FTIMESTP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTipos_bote_ex_mr_equipos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TIPOS_BOTE_EX_MR_EQUIPOS " & _
                   "  VALUES (" & _
                        TIPO_BOTE_EX_ID & "," & EQUIPO_ID & ",CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Insertar = TIPO_BOTE_EX_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTipos_bote_ex_mr_equipos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TIPOS_BOTE_EX_MR_EQUIPOS SET " & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " FTIMESTP = '" & FTIMESTP & "'" & _
                " WHERE TIPO_BOTE_EX_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTipos_bote_ex_mr_equipos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_BOTE_EX_MR_EQUIPOS " & _
                "    WHERE TIPO_BOTE_EX_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTipos_bote_ex_mr_equipos)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT E.ID_EQUIPO, E.NOMBRE, E.SERIE FROM TIPOS_BOTE_EX_MR_EQUIPOS A, EQUIPOS E WHERE A.EQUIPO_ID = E.ID_EQUIPO AND A.TIPO_BOTE_EX_ID = " & ID
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT TIPO_BOTE_EX_ID,EQUIPO_ID FROM TIPOS_BOTE_EX_MR_EQUIPOS ORDER BY EQUIPO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
