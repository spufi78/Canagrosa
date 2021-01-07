VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRex_botes_certificados_conf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsRex_botes_certificados_conf
'***************************************************************
Private BOTE_EX_ID As Long
Private ORDEN As Long
Private VALOR As String
Private TOLERANCIA As String
Private UNIDADES As String
Private CONFORME As String
'Mxxxx-I
Private CUSERID As Long
'Mxxxx-F
Private FECHA_CERTIFICACION As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsRex_botes_certificados_conf
'***************************************************************
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setVALOR(ByVal dato As String)
        VALOR = dato
End Property
Public Property Let setTOLERANCIA(ByVal dato As String)
        TOLERANCIA = dato
End Property
Public Property Let setUNIDADES(ByVal dato As String)
        UNIDADES = dato
End Property
Public Property Let setCONFORME(ByVal dato As String)
        CONFORME = dato
End Property
'Mxxxx-I
Public Property Let setCUSERID(ByVal dato As Long)
        CUSERID = dato
End Property
'Mxxxx-F
Public Property Let setFECHA_CERTIFICACION(ByVal dato As String)
        FECHA_CERTIFICACION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsRex_botes_certificados_conf
'***************************************************************
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getVALOR() As String
        getVALOR = VALOR
End Property
Public Property Get getTOLERANCIA() As String
        getTOLERANCIA = TOLERANCIA
End Property
Public Property Get getUNIDADES() As String
        getUNIDADES = UNIDADES
End Property
Public Property Get getCONFORME() As String
        getCONFORME = CONFORME
End Property
'Mxxxx-I
Public Property Get getCUSERID() As Long
        getCUSERID = CUSERID
End Property
'Mxxxx-F
Public Property Get getFECHA_CERTIFICACION() As String
        getFECHA_CERTIFICACION = FECHA_CERTIFICACION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsRex_botes_certificados_conf
'***************************************************************
Public Function Carga(ID As Long, ord As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REX_BOTES_CERTIFICADOS_CONFORMIDAD WHERE BOTE_EX_ID = " & ID & " AND ORDEN = " & ord
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                BOTE_EX_ID = rs("BOTE_EX_ID")
                ORDEN = rs("ORDEN")
                VALOR = rs("VALOR")
                TOLERANCIA = rs("TOLERANCIA")
                UNIDADES = rs("UNIDADES")
                CONFORME = rs("CONFORME")
                CUSERID = rs("CUSERID")
                FECHA_CERTIFICACION = rs("FECHA_CERTIFICACION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRex_botes_certificados_conf)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO REX_BOTES_CERTIFICADOS_CONFORMIDAD " & _
                   "  VALUES (" & _
                        BOTE_EX_ID & "," & ORDEN & ",'" & VALOR & "','" & TOLERANCIA & "','" & UNIDADES & "','" & CONFORME & "'," & USUARIO.getID_EMPLEADO & ",CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Insertar = BOTE_EX_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRex_botes_certificados_conf)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE REX_BOTES_CERTIFICADOS_CONFORMIDAD SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " VALOR = '" & VALOR & "'," & _
                        " TOLERANCIA = '" & TOLERANCIA & "'," & _
                        " UNIDADES = '" & UNIDADES & "'," & _
                        " CONFORME = '" & CONFORME & "'," & _
                        " CUSERID = " & USUARIO.getID_EMPLEADO & "," & _
                        " FECHA_CERTIFICACION = CURRENT_TIMESTAMP" & _
                " WHERE BOTE_EX_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRex_botes_certificados_conf)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM REX_BOTES_CERTIFICADOS_CONFORMIDAD " & _
                "    WHERE BOTE_EX_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRex_botes_certificados_conf)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REX_BOTES_CERTIFICADOS_CONFORMIDAD ORDER BY BOTE_EX_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoBote(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REX_BOTES_CERTIFICADOS_CONFORMIDAD WHERE BOTE_EX_ID = " & ID
        Set ListadoBote = datos_bd(consulta)
End Function
