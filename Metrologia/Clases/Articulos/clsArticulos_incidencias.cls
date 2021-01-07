VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticulos_incidencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSARTICULOS_INCIDENCIAS
'***************************************************************
Private ID_INCIDENCIA As Long
Private fecha As String
Private CODIGO As String
Private NOMBRE As String
Private DESCRIPCION As String
Private KM As Long
Private COSTE As String
Private COMENTARIO As String
Private ARTICULO_ID As Long
Private BATEA_ID As Long
Private CAMION_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSARTICULOS_INCIDENCIAS
'***************************************************************
Public Property Let setID_INCIDENCIA(ByVal dato As Long)
        ID_INCIDENCIA = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setKM(ByVal dato As Long)
        KM = dato
End Property
Public Property Let setCOSTE(ByVal dato As String)
        COSTE = dato
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
        COMENTARIO = dato
End Property
Public Property Let setARTICULO_ID(ByVal dato As Long)
        ARTICULO_ID = dato
End Property
Public Property Let setBATEA_ID(ByVal dato As Long)
        BATEA_ID = dato
End Property
Public Property Let setCAMION_ID(ByVal dato As Long)
        CAMION_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSARTICULOS_INCIDENCIAS
'***************************************************************
Public Property Get getID_INCIDENCIA() As Long
        getID_INCIDENCIA = ID_INCIDENCIA
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getKM() As Long
        getKM = KM
End Property
Public Property Get getCOSTE() As String
        getCOSTE = COSTE
End Property
Public Property Get getCOMENTARIO() As String
        getCOMENTARIO = COMENTARIO
End Property
Public Property Get getARTICULO_ID() As Long
        getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getBATEA_ID() As Long
        getBATEA_ID = BATEA_ID
End Property
Public Property Get getCAMION_ID() As Long
        getCAMION_ID = CAMION_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSARTICULOS_INCIDENCIAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS_INCIDENCIAS WHERE ID_INCIDENCIA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_INCIDENCIA = rs("ID_INCIDENCIA")
                fecha = rs("FECHA")
                CODIGO = rs("CODIGO")
                NOMBRE = rs("NOMBRE")
                DESCRIPCION = rs("DESCRIPCION")
                KM = rs("KM")
                COSTE = rs("COSTE")
                COMENTARIO = rs("COMENTARIO")
                CAMION_ID = rs("CAMION_ID")
                BATEA_ID = rs("BATEA_ID")
                ARTICULO_ID = rs("ARTICULO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsArticulos_incidencias)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_INCIDENCIA) FROM ARTICULOS_INCIDENCIAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_INCIDENCIA = 1
        Else
                ID_INCIDENCIA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO ARTICULOS_INCIDENCIAS " & _
                   "  VALUES (" & _
                        ID_INCIDENCIA & "," & "'" & fecha & "'" & "," & "'" & CODIGO & "'" & "," & _
                        "'" & NOMBRE & "'" & "," & "'" & DESCRIPCION & "'" & "," & KM & "," & _
                        "'" & COSTE & "'" & "," & "'" & COMENTARIO & "'" & "," & ARTICULO_ID & "," & BATEA_ID & "," & CAMION_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_INCIDENCIA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsArticulos_incidencias)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ARTICULOS_INCIDENCIAS SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " KM = " & KM & "," & _
                        " COSTE = '" & COSTE & "'," & _
                        " COMENTARIO = '" & COMENTARIO & "'," & _
                        " ARTICULO_ID = " & ARTICULO_ID & "," & _
                        " BATEA_ID = " & BATEA_ID & "," & _
                        " CAMION_ID = " & CAMION_ID & "" & _
                " WHERE ID_INCIDENCIA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsArticulos_incidencias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ARTICULOS_INCIDENCIAS " & _
                "    WHERE ID_INCIDENCIA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsArticulos_incidencias)", vbCritical, Err.Description
End Function
Public Function Listado(camion As Long, batea As Long) As ADODB.Recordset
        Dim consulta As String
        If camion <> 0 Then
            consulta = "SELECT * FROM ARTICULOS_INCIDENCIAS " & _
                       " WHERE CAMION_ID = " & camion & _
                       " ORDER BY ID_INCIDENCIA DESC"
        Else
            consulta = "SELECT * FROM ARTICULOS_INCIDENCIAS " & _
                       " WHERE BATEA_ID = " & batea & _
                       " ORDER BY ID_INCIDENCIA DESC"
        
        End If
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_INCIDENCIA,FECHA FROM ARTICULOS_INCIDENCIAS ORDER BY FECHA"
        Set Listado_Combo = datos_bd(consulta)
End Function
