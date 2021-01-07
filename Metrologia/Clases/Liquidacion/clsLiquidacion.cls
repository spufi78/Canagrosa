VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSLIQUIDACION
'***************************************************************
Private ID_LIQUIDACION As Long
Private DESCRIPCION As String
Private fdesde As String
Private fhasta As String
Private fliquidacion As String
Private AGENTE_ID As Long
Private USUARIO_ID As Long
Private TIMESTAMP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSLIQUIDACION
'***************************************************************
Public Property Let setID_LIQUIDACION(ByVal dato As Long)
        ID_LIQUIDACION = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFDESDE(ByVal dato As String)
        fdesde = dato
End Property
Public Property Let setFHASTA(ByVal dato As String)
        fhasta = dato
End Property
Public Property Let setFLIQUIDACION(ByVal dato As String)
        fliquidacion = dato
End Property
Public Property Let setAGENTE_ID(ByVal dato As Long)
        AGENTE_ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTIMESTAMP(ByVal dato As String)
        TIMESTAMP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSLIQUIDACION
'***************************************************************
Public Property Get getID_LIQUIDACION() As Long
        getID_LIQUIDACION = ID_LIQUIDACION
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFDESDE() As String
        getFDESDE = fdesde
End Property
Public Property Get getFHASTA() As String
        getFHASTA = fhasta
End Property
Public Property Get getFLIQUIDACION() As String
        getFLIQUIDACION = fliquidacion
End Property
Public Property Get getAGENTE_ID() As Long
        getAGENTE_ID = AGENTE_ID
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTIMESTAMP() As String
        getTIMESTAMP = TIMESTAMP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSLIQUIDACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM LIQUIDACION WHERE ID_LIQUIDACION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_LIQUIDACION = rs("ID_LIQUIDACION")
                DESCRIPCION = rs("DESCRIPCION")
                fdesde = rs("FDESDE")
                fhasta = rs("FHASTA")
                fliquidacion = rs("FLIQUIDACION")
                AGENTE_ID = rs("AGENTE_ID")
                USUARIO_ID = rs("USUARIO_ID")
                TIMESTAMP = rs("TIMESTAMP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsLiquidacion)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_LIQUIDACION) FROM LIQUIDACION"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_LIQUIDACION = 1
        Else
                ID_LIQUIDACION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO LIQUIDACION " & _
                   "  VALUES (" & _
                        ID_LIQUIDACION & "," & "'" & DESCRIPCION & "'" & "," & "'" & fdesde & "'" & "," & _
                        "'" & fhasta & "'" & "," & "'" & fliquidacion & "'" & "," & AGENTE_ID & "," & _
                        USUARIO_ID & "," & "current_timestamp" & _
                ")"
        execute_bd consulta
        Insertar = ID_LIQUIDACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsLiquidacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE LIQUIDACION SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FDESDE = '" & fdesde & "'," & _
                        " FHASTA = '" & fhasta & "'," & _
                        " FLIQUIDACION = '" & fliquidacion & "'," & _
                        " AGENTE_ID = " & AGENTE_ID & "," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TIMESTAMP = current_timestamp " & _
                " WHERE ID_LIQUIDACION = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsLiquidacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "DELETE FROM LIQUIDACION_DOCUMENTOS WHERE LIQUIDACION_ID = " & ID
        execute_bd "DELETE FROM LIQUIDACION WHERE ID_LIQUIDACION = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsLiquidacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM LIQUIDACION ORDER BY ID_LIQUIDACION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_LIQUIDACION,DESCRIPCION FROM LIQUIDACION ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
