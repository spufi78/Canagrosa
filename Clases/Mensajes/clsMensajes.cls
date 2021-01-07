VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMENSAJES
'***************************************************************
Private ID_MENSAJE As Long
Private ASUNTO As String
Private texto As String
Private ACCION As String
Private EMPLEADO_ID As Long
Private fecha_inicio As String
Private fecha_fin As String
Private HORA_INICIO As String
Private HORA_FIN As String
Private categoria As Integer
Private DURACION As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMENSAJES
'***************************************************************
Public Property Let setID_MENSAJE(ByVal dato As Long)
        ID_MENSAJE = dato
End Property
Public Property Let setASUNTO(ByVal dato As String)
        ASUNTO = dato
End Property
Public Property Let setTEXTO(ByVal dato As String)
        texto = dato
End Property
Public Property Let setACCION(ByVal dato As String)
        ACCION = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setFECHA_INICIO(ByVal dato As String)
        fecha_inicio = dato
End Property
Public Property Let setFECHA_FIN(ByVal dato As String)
        fecha_fin = dato
End Property
Public Property Let setHORA_INICIO(ByVal dato As String)
        HORA_INICIO = dato
End Property
Public Property Let setHORA_FIN(ByVal dato As String)
        HORA_FIN = dato
End Property
Public Property Let setCATEGORIA(ByVal dato As Integer)
        categoria = dato
End Property
Public Property Let setDURACION(ByVal dato As Integer)
        DURACION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMENSAJES
'***************************************************************
Public Property Get getID_MENSAJE() As Long
        getID_MENSAJE = ID_MENSAJE
End Property
Public Property Get getASUNTO() As String
        getASUNTO = ASUNTO
End Property
Public Property Get getTEXTO() As String
        getTEXTO = texto
End Property
Public Property Get getACCION() As String
        getACCION = ACCION
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getFECHA_INICIO() As String
        getFECHA_INICIO = fecha_inicio
End Property
Public Property Get getFECHA_FIN() As String
        getFECHA_FIN = fecha_fin
End Property
Public Property Get getHORA_INICIO() As String
        getHORA_INICIO = HORA_INICIO
End Property
Public Property Get getHORA_FIN() As String
        getHORA_FIN = HORA_FIN
End Property
Public Property Get getCATEGORIA() As Integer
        getCATEGORIA = categoria
End Property
Public Property Get getDURACION() As Integer
        getDURACION = DURACION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMENSAJES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM mensajes WHERE ID_MENSAJE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_MENSAJE = rs("ID_MENSAJE")
                ASUNTO = rs("ASUNTO")
                texto = rs("TEXTO")
                ACCION = rs("ACCION")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                fecha_inicio = rs("FECHA_INICIO")
                fecha_fin = rs("FECHA_FIN")
                HORA_INICIO = rs("HORA_INICIO")
                HORA_FIN = rs("HORA_FIN")
                categoria = rs("CATEGORIA")
                DURACION = rs("DURACION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMensajes)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_MENSAJE) FROM MENSAJES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_MENSAJE = 1
        Else
                ID_MENSAJE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO mensajes " & _
                   "  VALUES (" & _
                        ID_MENSAJE & "," & "'" & ASUNTO & "'" & "," & "'" & texto & "','" & ACCION & "'," & _
                        EMPLEADO_ID & "," & "'" & fecha_inicio & "'" & "," & "'" & CStr(fecha_fin) & "','" & _
                        HORA_INICIO & "','" & HORA_FIN & "'," & categoria & "," & DURACION & _
                ")"
        execute_bd consulta
        Insertar = ID_MENSAJE
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMensajes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE mensajes SET " & _
                        " ASUNTO = '" & ASUNTO & "'," & _
                        " TEXTO = '" & texto & "'," & _
                        " ACCION = '" & ACCION & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " FECHA_INICIO = '" & fecha_inicio & "'," & _
                        " FECHA_FIN = '" & fecha_fin & "'," & _
                        " HORA_INICIO = '" & HORA_INICIO & "'," & _
                        " HORA_FIN = '" & HORA_FIN & "'," & _
                        " CATEGORIA = " & categoria & "," & _
                        " DURACION = " & DURACION & _
                " WHERE ID_MENSAJE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMensajes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM mensajes " & _
                "    WHERE ID_MENSAJE = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMensajes)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT M.ID_MENSAJE,M.ASUNTO,MU.LEIDO,M.ACCION,M.TEXTO,M.FECHA_INICIO,M.FECHA_FIN " & _
                   " FROM mensajes M, mensajes_usuarios MU " & _
                   " WHERE M.ID_MENSAJE = MU.MENSAJE_ID " & _
                   "   AND MU.EMPLEADO_ID = " & USUARIO.getID_EMPLEADO & _
                   "   AND M.FECHA_INICIO <='" & Format(Date, "yyyy-mm-dd") & "' " & _
                   "   AND M.FECHA_FIN >='" & Format(Date, "yyyy-mm-dd") & "' " & _
                   " ORDER BY M.ID_MENSAJE"
        Set Listado = datos_bd(consulta, True)
End Function
Public Function Listado_completo(fecha As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT M.ID_MENSAJE,M.ASUNTO,MU.LEIDO,M.ACCION,M.TEXTO,M.FECHA_INICIO,M.FECHA_FIN,M.TEXTO, " & _
                   "       M.HORA_INICIO,M.HORA_FIN,M.CATEGORIA,M.DURACION " & _
                   " FROM mensajes M, mensajes_usuarios MU " & _
                   " WHERE M.ID_MENSAJE = MU.MENSAJE_ID " & _
                   "   AND M.FECHA_INICIO >='" & Format(fecha, "yyyy-mm-dd") & "' " & _
                   "   AND MU.EMPLEADO_ID = " & USUARIO.getID_EMPLEADO & _
                   " ORDER BY M.ID_MENSAJE DESC"
        Set Listado_completo = datos_bd(consulta, True)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_MENSAJE,ASUNTO FROM mensajes ORDER BY ASUNTO"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Mensajes_Sin_Leer() As Boolean
'        On Error Resume Next
        Dim consulta As String
        Dim rs As ADODB.Recordset
   On Error GoTo Mensajes_Sin_Leer_Error

        consulta = "SELECT M.ID_MENSAJE,M.ASUNTO,MU.LEIDO " & _
                   " FROM mensajes M, mensajes_usuarios MU " & _
                   " WHERE M.ID_MENSAJE = MU.MENSAJE_ID " & _
                   "   AND MU.EMPLEADO_ID = " & USUARIO.getID_EMPLEADO & _
                   "   AND M.FECHA_INICIO <='" & Format(Date, "yyyy-mm-dd") & "' " & _
                   "   AND M.FECHA_FIN >='" & Format(Date, "yyyy-mm-dd") & "' " & _
                   "   AND MU.LEIDO = 0 "
        Set rs = datos_bd(consulta, True)
        If rs.RecordCount = 0 Then
            Mensajes_Sin_Leer = False
        Else
            Mensajes_Sin_Leer = True
        End If

   On Error GoTo 0
   Exit Function

Mensajes_Sin_Leer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Mensajes_Sin_Leer of Módulo de clase clsMensajes"
End Function


