VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipoAccesorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsEquipoAccesorios
'***************************************************************
Private ID_ACCESORIO As Long
Private nombre As String
Private EQUIPO_ID As Long
Private ORDEN As Integer
Private CUSERID As String
Private FTIMESTP As String
Private FECHA_ALTA As String
Private FECHA_BAJA As String
Private EN_USO As Integer
Private ID_AUX As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsEquipoAccesorios
'***************************************************************
Public Property Let setID_ACCESORIO(ByVal dato As Long)
        ID_ACCESORIO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
Public Property Let setCUSERID(ByVal dato As String)
        CUSERID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
Public Property Let setFECHA_ALTA(ByVal dato As String)
        FECHA_ALTA = dato
End Property
Public Property Let setFECHA_BAJA(ByVal dato As String)
        FECHA_BAJA = dato
End Property
Public Property Let setEN_USO(ByVal dato As Integer)
        EN_USO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsEquipoAccesorios
'***************************************************************
Public Property Get getID_ACCESORIO() As Long
        getID_ACCESORIO = ID_ACCESORIO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
Public Property Get getCUSERID() As String
        getCUSERID = CUSERID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
Public Property Get getFECHA_ALTA() As String
        getFECHA_ALTA = FECHA_ALTA
End Property
Public Property Get getFECHA_BAJA() As String
        getFECHA_BAJA = FECHA_BAJA
End Property
Public Property Get getEN_USO() As Integer
        getEN_USO = EN_USO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsEquipoAccesorios
'***************************************************************
Public Function Carga(ID As Long, ID_EQUIPO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        
        consulta = " select eq_accesorios.*, Equipos.nombre from eq_accesorios "
        consulta = consulta & " inner join equipos on eq_accesorios.id_accesorio = equipos.ID_EQUIPO "
        consulta = consulta & " WHERE eq_accesorios.equipo_id =  " & CStr(ID_EQUIPO)
        consulta = consulta & " AND eq_accesorios.ID_ACCESORIO = " & CStr(id)
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ACCESORIO = rs("ID_ACCESORIO")
                nombre = rs("NOMBRE")
                EQUIPO_ID = rs("EQUIPO_ID")
                ORDEN = rs("ORDEN")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                If Not IsNull(rs("FECHA_ALTA")) Then
                    FECHA_ALTA = rs("FECHA_ALTA")
                Else
                    FECHA_ALTA = Now
                End If
                FECHA_BAJA = rs("FECHA_BAJA")
                EN_USO = rs("EN_USO")
                ID_AUX = 1
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipoAccesorios)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        
        
        'ATENCION, AQUI RECOGE EL ORDEN, NO EL ID_ACCESORIO (QUE ES EL DEL EQUIPO QUE FUNCIONA COMO ACCESORIO)
        
        consulta = "select coalesce(count(orden), 0) + 1  from eq_accesorios where equipo_id = " & EQUIPO_ID
        Set rs = datos_bd(consulta)
        rs.MoveFirst
        ORDEN = CInt(rs(0))
        
        
        consulta = "INSERT INTO eq_accesorios (ID_ACCESORIO, EQUIPO_ID, ORDEN,CUSERID, FTIMESTP, FECHA_ALTA, FECHA_BAJA, EN_USO) " & _
                   "  VALUES (" & _
                        ID_ACCESORIO & "," & EQUIPO_ID & "," & _
                        ORDEN & "," & "'" & CUSERID & "'" & ", localtimestamp," & _
                        "'" & Format(FECHA_ALTA, "YYYY-MM-DD Hh:Nn") & "'" & "," & "'" & Format(FECHA_BAJA, "YYYY-MM-DD Hh:Nn") & "'" & "," & EN_USO & _
                ")"
        execute_bd consulta
        Insertar = ID_ACCESORIO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipoAccesorios)", vbCritical, Err.Description
End Function
Public Function Modificar() As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE eq_accesorios SET " & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " ORDEN = " & ORDEN & "," & _
                        " CUSERID = '" & USUARIO.getUSUARIO & "'," & _
                        " FTIMESTP = LOCALTIMESTAMP," & _
                        " FECHA_ALTA = '" & Format(FECHA_ALTA, "YYYY-MM-DD Hh:Nn") & "'," & _
                        " FECHA_BAJA = '" & Format(FECHA_BAJA, "YYYY-MM-DD Hh:Nn") & "'," & _
                        " EN_USO = " & EN_USO & "" & _
                " WHERE ID_ACCESORIO = " & ID_ACCESORIO & _
                " AND equipo_id = " & EQUIPO_ID
                
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipoAccesorios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID_ACC As Long, id_eq As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM eq_accesorios " & _
                "    WHERE ID_ACCESORIO = " & ID_ACC & _
                "    AND EQUIPO_ID = " & id_eq
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipoAccesorios)", vbCritical, Err.Description
End Function
Public Function Listado(ByVal ID_EQUIPO As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = " select eq_accesorios.*, Equipos.nombre from eq_accesorios "
        consulta = consulta & " inner join equipos on eq_accesorios.id_accesorio = equipos.id_equipo "
        consulta = consulta & " WHERE eq_accesorios.equipo_id =  " & CStr(ID_EQUIPO)
        consulta = consulta & " ORDER BY eq_accesorios.ID_ACCESORIO "
        Set Listado = datos_bd(consulta)
End Function

Public Property Get getID_AUX() As Integer

    getID_AUX = ID_AUX

End Property

Public Property Let setID_AUX(ByVal intID_AUX As Integer)

    ID_AUX = intID_AUX

End Property
