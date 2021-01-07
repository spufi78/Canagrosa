VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_bote_ex_mr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_BOTE_EX_MR
'***************************************************************
Private ID_TIPO_BOTE_EX_MR As Long
Private TIPO_BOTE_EX_ID As Long
Private USO_PREVISTO As String
Private PERIODICIDAD As String
Private EQUIPO_ID As Long
Private PNTS As String
Private ENSAYO_ENAC As Long
Private OBSERVACIONES As String
Private CUSERID As Long
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_BOTE_EX_MR
'***************************************************************
Public Property Let setID_TIPO_BOTE_EX_MR(ByVal dato As Long)
        ID_TIPO_BOTE_EX_MR = dato
End Property
Public Property Let setTIPO_BOTE_EX_ID(ByVal dato As Long)
        TIPO_BOTE_EX_ID = dato
End Property
Public Property Let setUSO_PREVISTO(ByVal dato As String)
        USO_PREVISTO = dato
End Property
Public Property Let setPERIODICIDAD(ByVal dato As String)
        PERIODICIDAD = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setPNTS(ByVal dato As String)
        PNTS = dato
End Property
Public Property Let setENSAYO_ENAC(ByVal dato As Long)
        ENSAYO_ENAC = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setCUSERID(ByVal dato As Long)
        CUSERID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_BOTE_EX_MR
'***************************************************************
Public Property Get getID_TIPO_BOTE_EX_MR() As Long
        getID_TIPO_BOTE_EX_MR = ID_TIPO_BOTE_EX_MR
End Property
Public Property Get getTIPO_BOTE_EX_ID() As Long
        getTIPO_BOTE_EX_ID = TIPO_BOTE_EX_ID
End Property
Public Property Get getUSO_PREVISTO() As String
        getUSO_PREVISTO = USO_PREVISTO
End Property
Public Property Get getPERIODICIDAD() As String
        getPERIODICIDAD = PERIODICIDAD
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getPNTS() As String
        getPNTS = PNTS
End Property
Public Property Get getENSAYO_ENAC() As Long
        getENSAYO_ENAC = ENSAYO_ENAC
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getCUSERID() As Long
        getCUSERID = CUSERID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTIPOS_BOTE_EX_MR
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_BOTE_EX_MR WHERE ID_TIPO_BOTE_EX_MR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TIPO_BOTE_EX_MR = rs("ID_TIPO_BOTE_EX_MR")
                TIPO_BOTE_EX_ID = rs("TIPO_BOTE_EX_ID")
                USO_PREVISTO = rs("USO_PREVISTO")
                PERIODICIDAD = rs("PERIODICIDAD")
                EQUIPO_ID = rs("EQUIPO_ID")
                PNTS = rs("PNTS")
                ENSAYO_ENAC = rs("ENSAYO_ENAC")
                OBSERVACIONES = rs("OBSERVACIONES")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(CLSTIPOS_BOTE_EX_MR)", vbCritical, Err.Description
End Function
Public Function CargaTipo(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_BOTE_EX_MR WHERE TIPO_BOTE_EX_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                CargaTipo = False
        Else
                ID_TIPO_BOTE_EX_MR = rs("ID_TIPO_BOTE_EX_MR")
                TIPO_BOTE_EX_ID = rs("TIPO_BOTE_EX_ID")
                USO_PREVISTO = rs("USO_PREVISTO")
                PERIODICIDAD = rs("PERIODICIDAD")
                EQUIPO_ID = rs("EQUIPO_ID")
                PNTS = rs("PNTS")
                ENSAYO_ENAC = rs("ENSAYO_ENAC")
                OBSERVACIONES = rs("OBSERVACIONES")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                CargaTipo = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaTipo = False
        MsgBox "Error al cargar los datos por tipo(CLSTIPOS_BOTE_EX_MR)", vbCritical, Err.Description
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_BOTE_EX_MR) FROM TIPOS_BOTE_EX_MR"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_BOTE_EX_MR = 1
    Else
        ID_TIPO_BOTE_EX_MR = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
        On Error GoTo fallo
        CrearID
        Dim consulta As String
        consulta = "INSERT INTO TIPOS_BOTE_EX_MR " & _
                   "  VALUES (" & _
                        ID_TIPO_BOTE_EX_MR & "," & TIPO_BOTE_EX_ID & "," & "'" & USO_PREVISTO & "'" & "," & _
                        "'" & PERIODICIDAD & "'" & "," & EQUIPO_ID & "," & "'" & PNTS & "'" & "," & _
                        ENSAYO_ENAC & "," & "'" & OBSERVACIONES & "'" & "," & USUARIO.getID_EMPLEADO & "," & _
                        "CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Insertar = ID_TIPO_BOTE_EX_MR
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (CLSTIPOS_BOTE_EX_MR)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TIPOS_BOTE_EX_MR SET " & _
                        " TIPO_BOTE_EX_ID = " & TIPO_BOTE_EX_ID & "," & _
                        " USO_PREVISTO = '" & USO_PREVISTO & "'," & _
                        " PERIODICIDAD = '" & PERIODICIDAD & "'," & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " PNTS = '" & PNTS & "'," & _
                        " ENSAYO_ENAC = " & ENSAYO_ENAC & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " CUSERID = " & CUSERID & "," & _
                        " FTIMESTP = '" & FTIMESTP & "'" & _
                " WHERE ID_TIPO_BOTE_EX_MR = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (CLSTIPOS_BOTE_EX_MR)", vbCritical, Err.Description
End Function
Public Function ModificarTipo(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TIPOS_BOTE_EX_MR SET " & _
                        " TIPO_BOTE_EX_ID = " & TIPO_BOTE_EX_ID & "," & _
                        " USO_PREVISTO = '" & USO_PREVISTO & "'," & _
                        " PERIODICIDAD = '" & PERIODICIDAD & "'," & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " PNTS = '" & PNTS & "'," & _
                        " ENSAYO_ENAC = " & ENSAYO_ENAC & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " CUSERID = " & CUSERID & "," & _
                        " FTIMESTP = '" & FTIMESTP & "'" & _
                " WHERE TIPO_BOTE_EX_ID = " & ID
        execute_bd consulta
        ModificarTipo = True
        Exit Function
fallo:
        ModificarTipo = False
        MsgBox "Error al Modificar Tipo(CLSTIPOS_BOTE_EX_MR)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_BOTE_EX_MR " & _
                "    WHERE ID_TIPO_BOTE_EX_MR = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (CLSTIPOS_BOTE_EX_MR)", vbCritical, Err.Description
End Function
Public Function EliminarTipo(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_BOTE_EX_MR " & _
                "    WHERE TIPO_BOTE_EX_ID = " & ID
        execute_bd consulta
        EliminarTipo = True
        Exit Function
fallo:
        EliminarTipo = False
        MsgBox "Error al EliminarTipo (CLSTIPOS_BOTE_EX_MR)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_BOTE_EX_MR ORDER BY ID_TIPO_BOTE_EX_MR"
        Set Listado = datos_bd(consulta)
End Function
