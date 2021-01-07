VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_Accesorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_ACCESORIOS
'***************************************************************
Private ID_ACCESORIO As Long
Private EQUIPO_ID As Long
Private nombre As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_ACCESORIOS
'***************************************************************
Public Property Let setID_ACCESORIO(ByVal dato As Long)
        ID_ACCESORIO = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
       nombre = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_ACCESORIOS
'***************************************************************
Public Property Get getID_ACCESORIO() As Long
        getID_ACCESORIO = ID_ACCESORIO
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_ACCESORIOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT eq_accesorios.id_accesorio, eq_accesorios.equipo_id, eq_accesorios.orden, eq_accesorios.cuserid, eq_accesorios.ftimestp, eq_accesorios.fecha_alta, eq_accesorios.fecha_baja, eq_accesorios.en_uso, coalesce(equipos.nombre,'N/A') as nombre  FROM eq_accesorios left outer join equipos on eq_accesorios.equipo_id = equipos.id_equipo WHERE eq_accesorios.EQUIPO_ID = " & id
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            EQUIPO_ID = rs("EQUIPO_ID")
            nombre = rs("NOMBRE")
            ORDEN = rs("ORDEN")
            Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_Accesorios)", vbCritical, Err.Description
End Function

Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_ACCESORIO) FROM eq_accesorios"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ACCESORIO = 1
        Else
                ID_ACCESORIO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        If Not duplicado Then
            If ID_ACCESORIO = 0 Then
                Me.CrearID
            End If
                
            consulta = "INSERT INTO eq_accesorios VALUES (" & _
                                    ID_ACCESORIO & ", " & "'" & nombre & "'" & ", " & EQUIPO_ID & ", " & ORDEN & ", " & _
                                    "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                       ")"
            execute_bd consulta
            Insertar = ID_ACCESORIO
        End If
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_Accesorios)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE eq_accesorios SET " & _
                   "       EQUIPO_ID = " & EQUIPO_ID & "" & _
                   "       NOMBRE = '" & nombre & "'," & _
                   "       ORDEN = " & ORDEN & _
                   "       CUSERID = '" & USUARIO.getUSUARIO & "'" & _
                   "       FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE EQUIPO_ID = " & ID
        execute_bd consulta
        Modificar = True
        
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos_Accesorios)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM eq_accesorios " & _
                   " WHERE ID_ACCESORIO = " & ID
        execute_bd consulta
        Eliminar = True
        
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_Accesorios)", vbCritical, Err.Description
End Function

' Listado de todos los accesorios
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT E.NOMBRE, ED.NOMBRE ACCESORIO " & _
                   "  FROM EQUIPOS E, eq_accesorios EA " & _
                   " WHERE E.ID_EQUIPO = EA.EQUIPO_ID " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function

Public Function lista_por_equipos(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT EA.NOMBRE, EA.ID_ACCESORIO " & _
                   "  FROM eq_accesorios EA " & _
                   " WHERE EQUIPO_ID = " & ID
        Set lista_por_equipos = datos_bd(consulta)
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT COUNT(*) AS TOTAL " & _
               "  FROM eq_accesorios " & _
               " WHERE EQUIPO_ID = " & EQUIPO_ID & "" & _
               "   AND NOMBRE = '" & nombre & "'"
    Set rs = datos_bd(consulta)
    If rs.Fields("TOTAL") = 0 Then
        duplicado = False
    Else
        MsgBox "Ese accesorio ya está asociado al equipo.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
