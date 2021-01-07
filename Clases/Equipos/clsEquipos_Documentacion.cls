VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_Documentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_DOCUMENTACION
'***************************************************************
Private ID_DOCUMENTO As Long
Private EQUIPO_ID As Long
Private RUTA_DOCUMENTO As String
Private NOMBRE_DOCUMENTO As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_DOCUMENTACION
'***************************************************************
Public Property Let setID_DOCUMENTO(ByVal dato As Long)
        ID_DOCUMENTO = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setRUTA_DOCUMENTO(ByVal dato As String)
        RUTA_DOCUMENTO = Replace(dato, "\", "/")
End Property
Public Property Let setNOMBRE_DOCUMENTO(ByVal dato As String)
        NOMBRE_DOCUMENTO = Replace(dato, "\", "/")
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_DOCUMENTACION
'***************************************************************
Public Property Get getID_DOCUMENTO() As Long
        getID_DOCUMENTO = ID_DOCUMENTO
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getRUTA_DOCUMENTO() As String
        getRUTA_DOCUMENTO = Replace(RUTA_DOCUMENTO, "/", "\")
End Property
Public Property Get getNOMBRE_DOCUMENTO() As String
        getNOMBRE_DOCUMENTO = Replace(NOMBRE_DOCUMENTO, "/", "\")
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_DOCUMENTACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_DOCUMENTOS WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            EQUIPO_ID = rs("EQUIPO_ID")
            RUTA_DOCUMENTO = rs("RUTA_DOCUMENTO")
            NOMBRE_DOCUMENTO = rs("NOMBRE_DOCUMENTO")
            ORDEN = rs("ORDEN")
            Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_Documentacion)", vbCritical, Err.Description
End Function

Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_DOCUMENTO) FROM EQUIPOS_DOCUMENTACION"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_DOCUMENTO = 1
        Else
                ID_DOCUMENTO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        If Not duplicado Then
            If ID_DOCUMENTO = 0 Then
                Me.CrearID
            End If
                
            consulta = "INSERT INTO EQUIPOS_DOCUMENTACION VALUES (" & _
                                    ID_DOCUMENTO & ", " & EQUIPO_ID & "," & "'" & RUTA_DOCUMENTO & "'" & "," & "'" & NOMBRE_DOCUMENTO & "'" & "," & ORDEN & ", " & _
                                    "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                       ")"
            execute_bd consulta
            Insertar = ID_DOCUMENTO
        End If
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_Documentacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_DOCUMENTACION SET " & _
                   "       EQUIPO_ID = " & EQUIPO_ID & ", " & _
                   "       RUTA_DOCUMENTO = '" & RUTA_DOCUMENTO & "', " & _
                   "       NOMBRE_DOCUMENTO = '" & NOMBRE_DOCUMENTO & "', " & _
                   "       ORDEN = " & ORDEN & ", " & _
                   "       CUSERID = '" & USUARIO.getUSUARIO & "', " & _
                   "       FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE EQUIPO_ID = " & ID
        execute_bd consulta
        Modificar = True
        
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos_Documentacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM EQUIPOS_DOCUMENTACION " & _
                   " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        Eliminar = True
        
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_Documentacion)", vbCritical, Err.Description
End Function
' Listado de todos los documentos
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT E.NOMBRE, ED.RUTA_DOCUMENTO, ED.NOMBRE_DOCUMENTO " & _
                   "  FROM EQUIPOS E, EQUIPOS_DOCUMENTACION ED" & _
                   " WHERE E.ID_EQUIPO = ED.EQUIPO_ID " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
Public Function lista_por_equipos(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ED.ID_DOCUMENTO, ED.RUTA_DOCUMENTO, ED.NOMBRE_DOCUMENTO " & _
                   "  FROM EQUIPOS_DOCUMENTACION ED" & _
                   " WHERE EQUIPO_ID = " & ID
        Set lista_por_equipos = datos_bd(consulta)
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT COUNT(*) AS TOTAL " & _
               "  FROM EQUIPOS_DOCUMENTACION " & _
               " WHERE EQUIPO_ID = " & EQUIPO_ID & "" & _
               "   AND RUTA_DOCUMENTO = '" & RUTA_DOCUMENTO & "'"
    Set rs = datos_bd(consulta)
    If rs.Fields("TOTAL") = 0 Then
        duplicado = False
    Else
        MsgBox "Ese documento ya está asociado al equipo.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
