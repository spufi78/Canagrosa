VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipoEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsEquipoEventos
'***************************************************************
Private ID_EVENTOEQUIPO As Long
Private EVENTO_ID As Long
Private RAZON_ID As Long
Private CVM_ID As Long
Private OBSERVACIONES As String
Private EQUIPO_ID As Long
Private USUARIO_ID As Long
Private TS As String
Private EVENTO As String
Private RAZON As String
Private EQUIPO As String

Private mvarstrVistaListado As String
Public Function Devolver_Listado_Razones_Evento_Combo(ByVal prmid_evento As Long) As ADODB.Recordset

    Dim strIds_razones As String
    Dim strCad As String
    
    strIds_razones = devolver_parametros_evento(prmid_evento)
    strCad = "SELECT 0 as ID, '' as RAZON"
    
    If Trim(strIds_razones) <> "" Then
        strCad = "SELECT VALOR as ID, DESCRIPCION as RAZON FROM decodificadora where codigo  = " & DECODIFICADORA.EQ_RAZONES_EVENTOS & " AND valor in (" & strIds_razones & ") order by valor"
    End If
    
    Set Devolver_Listado_Razones_Evento_Combo = datos_bd(strCad)
    
    
End Function

Public Function devolver_parametros_evento(ByVal prmid_evento As Long) As String

On Error GoTo devolver_parametros_evento_Error

    devolver_parametros_evento = ""

    Dim rs As ADODB.Recordset
    
    Set rs = datos_bd("SELECT * FROM decodificadora where codigo = " & DECODIFICADORA.EQ_EVENTOS & " AND valor = " & CStr(prmid_evento))
    
    If Not rs Is Nothing Then
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            devolver_parametros_evento = rs("parametros")
        End If
    End If
        

On Error GoTo 0
    Exit Function
devolver_parametros_evento_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure devolver_parametros_evento of Módulo de clase clsEquipoEventos"
    devolver_parametros_evento = ""
End Function

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsEquipoEventos
'***************************************************************
Public Property Let setID_EVENTOEQUIPO(ByVal dato As Long)
        ID_EVENTOEQUIPO = dato
End Property
Public Property Let setEVENTO_ID(ByVal dato As Long)
        EVENTO_ID = dato
End Property
Public Property Let setRAZON_ID(ByVal dato As Long)
        RAZON_ID = dato
End Property
Public Property Let setCVM_ID(ByVal dato As Long)
        CVM_ID = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
Public Property Let setEVENTO(ByVal dato As String)
        EVENTO = dato
End Property
Public Property Let setRAZON(ByVal dato As String)
        RAZON = dato
End Property
Public Property Let setEQUIPO(ByVal dato As String)
        EQUIPO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsEquipoEventos
'***************************************************************
Public Property Get getID_EVENTOEQUIPO() As Long
        getID_EVENTOEQUIPO = ID_EVENTOEQUIPO
End Property
Public Property Get getEVENTO_ID() As Long
        getEVENTO_ID = EVENTO_ID
End Property
Public Property Get getRAZON_ID() As Long
        getRAZON_ID = RAZON_ID
End Property
Public Property Get getCVM_ID() As Long
        getCVM_ID = CVM_ID
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
Public Property Get getEVENTO() As String
        getEVENTO = EVENTO
End Property
Public Property Get getRAZON() As String
        getRAZON = RAZON
End Property
Public Property Get getEQUIPO() As String
        getEQUIPO = EQUIPO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsEquipoEventos
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = mvarstrVistaListado & " WHERE EQ_EVENTOS.ID_EVENTOEQUIPO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_EVENTOEQUIPO = rs("ID_EVENTOEQUIPO")
                EVENTO_ID = rs("EVENTO_ID")
                RAZON_ID = rs("RAZON_ID")
                CVM_ID = rs("CVM_ID")
                OBSERVACIONES = rs("OBSERVACIONES")
                EQUIPO_ID = rs("EQUIPO_ID")
                USUARIO_ID = rs("USUARIO_ID")
                TS = rs("TS")
                EVENTO = rs("EVENTO")
                RAZON = rs("RAZON")
                EQUIPO = rs("EQUIPO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipoEventos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_EVENTOEQUIPO) FROM EQ_EVENTOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_EVENTOEQUIPO = 1
        Else
                ID_EVENTOEQUIPO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO EQ_EVENTOS (ID_EVENTOEQUIPO, TS)" & _
                   "  VALUES (" & ID_EVENTOEQUIPO & ", LOCALTIMESTAMP)"
        execute_bd consulta
        
        Insertar = ID_EVENTOEQUIPO
        
        Modificar ID_EVENTOEQUIPO
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipoEventos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQ_EVENTOS SET " & _
                        " EVENTO_ID = " & EVENTO_ID & "," & _
                        " RAZON_ID = " & RAZON_ID & "," & _
                        " CVM_ID = " & CVM_ID & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TS = LOCALTIMESTAMP" & _
                " WHERE ID_EVENTOEQUIPO = " & ID
        
        execute_bd consulta
        Modificar = True
                
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipoEventos)", vbCritical, Err.Description
End Function


Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EQ_EVENTOS " & _
                "    WHERE ID_EVENTOEQUIPO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipoEventos)", vbCritical, Err.Description
End Function
Public Function Listado(Optional ByVal prmid_equipo As Long = 0) As ADODB.Recordset
        Dim consulta As String
        consulta = mvarstrVistaListado
        consulta = consulta & " WHERE 1=1 "
        If prmid_equipo <> 0 Then
            consulta = consulta & " AND EQ_EVENTOS.EQUIPO_ID = " & CStr(prmid_equipo)
        End If
        consulta = consulta & " ORDER BY EQ_EVENTOS.TS DESC"
        
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_EVENTOEQUIPO,EVENTO_ID FROM EQ_EVENTOS ORDER BY EVENTO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function

Private Sub Class_Initialize()

mvarstrVistaListado = "SELECT EQ_EVENTOS.*, EQUIPOS.NOMBRE AS EQUIPO, CONCAT(USUARIOS.NOMBRE, ' ', USUARIOS.APELLIDOS) AS USUARIO, DECODIFICADORA_EVENTOS.DESCRIPCION AS EVENTO, DECODIFICADORA_RAZONES.DESCRIPCION AS RAZON "
mvarstrVistaListado = mvarstrVistaListado & " FROM EQ_EVENTOS "
mvarstrVistaListado = mvarstrVistaListado & " INNER JOIN EQUIPOS ON EQ_EVENTOS.EQUIPO_ID = EQUIPOS.ID_EQUIPO "
mvarstrVistaListado = mvarstrVistaListado & " INNER JOIN usuarios ON EQ_EVENTOS.USUARIO_ID = USUARIOS.ID_EMPLEADO "
mvarstrVistaListado = mvarstrVistaListado & " LEFT OUTER JOIN decodificadora AS DECODIFICADORA_EVENTOS ON EQ_EVENTOS.EVENTO_ID = DECODIFICADORA_EVENTOS.VALOR AND DECODIFICADORA_EVENTOS.CODIGO=" & DECODIFICADORA.EQ_EVENTOS & " "
mvarstrVistaListado = mvarstrVistaListado & " LEFT OUTER JOIN decodificadora AS DECODIFICADORA_RAZONES ON EQ_EVENTOS.RAZON_ID = DECODIFICADORA_RAZONES.VALOR AND DECODIFICADORA_RAZONES.CODIGO=" & DECODIFICADORA.EQ_RAZONES_EVENTOS & " "

End Sub


