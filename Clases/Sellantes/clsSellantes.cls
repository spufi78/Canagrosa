VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSellantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSELLANTES
'***************************************************************
Private ID_SELLANTE As Long
Private ENSAYO As String
Private ENSAYO_INGLES As String
Private CLIENTE_ID As Long
Private PROCESO As String
Private PROCESO_INGLES As String
Private INSTALACION As String
Private INSTALACION_INGLES As String
Private PREPARACION As String
Private PREPARACION_INGLES As String
Private OBSERVACIONES As String
Private OBSERVACIONES_DE As String
Private producto As String
Private ENAC As Integer
Private anulado As Integer

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSELLANTES
'***************************************************************
Public Property Let setID_SELLANTE(ByVal dato As Long)
        ID_SELLANTE = dato
End Property
Public Property Let setENSAYO(ByVal dato As String)
        ENSAYO = dato
End Property
Public Property Let setENSAYO_INGLES(ByVal dato As String)
        ENSAYO_INGLES = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property

Public Property Let setPROCESO(ByVal dato As String)
        PROCESO = dato
End Property
Public Property Let setPROCESO_INGLES(ByVal dato As String)
        PROCESO_INGLES = dato
End Property
Public Property Let setINSTALACION(ByVal dato As String)
        INSTALACION = dato
End Property
Public Property Let setINSTALACION_INGLES(ByVal dato As String)
        INSTALACION_INGLES = dato
End Property
Public Property Let setPREPARACION(ByVal dato As String)
        PREPARACION = dato
End Property
Public Property Let setPREPARACION_INGLES(ByVal dato As String)
        PREPARACION_INGLES = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setOBSERVACIONES_DE(ByVal dato As String)
        OBSERVACIONES_DE = dato
End Property
Public Property Let setPRODUCTO(ByVal dato As String)
        producto = dato
End Property
Public Property Let setENAC(ByVal dato As Integer)
        ENAC = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
        anulado = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSELLANTES
'***************************************************************
Public Property Get getID_SELLANTE() As Long
        getID_SELLANTE = ID_SELLANTE
End Property
Public Property Get getENSAYO() As String
        getENSAYO = ENSAYO
End Property
Public Property Get getENSAYO_INGLES() As String
        getENSAYO_INGLES = ENSAYO_INGLES
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property

Public Property Get getPROCESO() As String
        getPROCESO = PROCESO
End Property
Public Property Get getPROCESO_INGLES() As String
        getPROCESO_INGLES = PROCESO_INGLES
End Property
Public Property Get getINSTALACION() As String
        getINSTALACION = INSTALACION
End Property
Public Property Get getINSTALACION_INGLES() As String
        getINSTALACION_INGLES = INSTALACION_INGLES
End Property
Public Property Get getPREPARACION() As String
        getPREPARACION = PREPARACION
End Property
Public Property Get getPREPARACION_INGLES() As String
        getPREPARACION_INGLES = PREPARACION_INGLES
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getOBSERVACIONES_DE() As String
        getOBSERVACIONES_DE = OBSERVACIONES_DE
End Property
Public Property Get getPRODUCTO() As String
        getPRODUCTO = producto
End Property
Public Property Get getENAC() As Integer
        getENAC = ENAC
End Property
Public Property Get getANULADO() As Integer
        getANULADO = anulado
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSELLANTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES WHERE ID_SELLANTE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SELLANTE = rs("ID_SELLANTE")
                ENSAYO = rs("ENSAYO")
                ENSAYO_INGLES = rs("ENSAYO_INGLES")
                CLIENTE_ID = rs("CLIENTE_ID")
                PROCESO = rs("PROCESO")
                PROCESO_INGLES = rs("PROCESO_INGLES")
                INSTALACION = rs("INSTALACION")
                INSTALACION_INGLES = rs("INSTALACION_INGLES")
                PREPARACION = rs("PREPARACION")
                PREPARACION_INGLES = rs("PREPARACION_INGLES")
                OBSERVACIONES = rs("OBSERVACIONES")
                OBSERVACIONES_DE = rs("OBSERVACIONES_DE")
                producto = rs("PRODUCTO")
                ENAC = rs("ENAC")
                anulado = rs("ANULADO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSellantes)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_SELLANTE) FROM SELLANTES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_SELLANTE = 1
        Else
                ID_SELLANTE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO SELLANTES " & _
                   "  VALUES (" & _
                        ID_SELLANTE & "," & "'" & ENSAYO & "'" & "," & "'" & ENSAYO_INGLES & "'" & "," & CLIENTE_ID & "," & _
                        "'" & PROCESO & "'" & "," & "'" & PROCESO_INGLES & "'" & "," & "'" & INSTALACION & "'" & "," & _
                        "'" & INSTALACION_INGLES & "'" & "," & "'" & PREPARACION & "'" & "," & "'" & PREPARACION_INGLES & "'" & "," & _
                        "'" & producto & "'," & ENAC & ",0,'" & OBSERVACIONES & "','" & OBSERVACIONES_DE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_SELLANTE
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSellantes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES SET " & _
                        " ENSAYO = '" & ENSAYO & "'," & _
                        " ENSAYO_INGLES = '" & ENSAYO_INGLES & "'," & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " PROCESO = '" & PROCESO & "'," & _
                        " PROCESO_INGLES = '" & PROCESO_INGLES & "'," & _
                        " INSTALACION = '" & INSTALACION & "'," & _
                        " INSTALACION_INGLES = '" & INSTALACION_INGLES & "'," & _
                        " PREPARACION = '" & PREPARACION & "'," & _
                        " PREPARACION_INGLES = '" & PREPARACION_INGLES & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " OBSERVACIONES_DE = '" & OBSERVACIONES_DE & "'," & _
                        " ENAC = " & ENAC & "," & _
                        " PRODUCTO = '" & producto & "'" & _
                " WHERE ID_SELLANTE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSellantes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("SELECT * FROM SELLANTES_RESULTADOS WHERE SELLANTE_ID = " & ID)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar el sellante, existen muestras que lo contienen.", vbExclamation, App.Title
            Exit Function
        End If
        execute_bd "DELETE FROM SELLANTES WHERE ID_SELLANTE = " & ID
        execute_bd "DELETE FROM SELLANTES_ENSAYOS WHERE SELLANTE_ID = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSellantes)", vbCritical, Err.Description
End Function
Public Function Anular(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "UPDATE SELLANTES SET ANULADO = " & USUARIO.getID_EMPLEADO & " WHERE ID_SELLANTE = " & ID
        Anular = True
        Exit Function
fallo:
        Anular = False
        MsgBox "Error al Anular (clsSellantes)", vbCritical, Err.Description
End Function
Public Function Listado(Lensayo As String, Lcliente As String, Lproceso As String, Lproducto As String, chkAnulado As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim s As String
        If Trim(Lensayo) <> "" Then
            s = s & " AND ENSAYO LIKE '%" & Lensayo & "%'"
        End If
        If Trim(Lproceso) <> "" Then
            s = s & " AND proceso LIKE '%" & Lproceso & "%'"
        End If
        If Trim(Lproducto) <> "" Then
            s = s & " AND producto LIKE '%" & Lproducto & "%'"
        End If
        If Trim(Lcliente) <> "" Then
            s = s & " AND S.CLIENTE_ID = " & Lcliente
        End If
        If chkAnulado = False Then
            s = s & " AND S.ANULADO = 0 "
        End If
        consulta = "SELECT S.ID_SELLANTE, S.ENSAYO,C.NOMBRE, S.PROCESO,S.PRODUCTO,S.ANULADO " & _
                   "  FROM SELLANTES S, CLIENTES C " & _
                   " WHERE S.CLIENTE_ID = C.ID_CLIENTE " & _
                   s & _
                   " ORDER BY S.ENSAYO, C.NOMBRE "
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT S.ID_SELLANTE,CONCAT(S.PRODUCTO,' (',C.NOMBRE,')')" & _
                   "  FROM SELLANTES S, CLIENTES C " & _
                   " WHERE S.CLIENTE_ID = C.ID_CLIENTE " & _
                   " ORDER BY S.PRODUCTO"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_Combo_Clientes() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT DISTINCT C.ID_CLIENTE AS C1,C.NOMBRE AS C2" & _
                   "  FROM SELLANTES S, CLIENTES C " & _
                   " WHERE S.CLIENTE_ID = C.ID_CLIENTE " & _
                   " ORDER BY C.NOMBRE"
        Set Listado_Combo_Clientes = datos_bd(consulta)
End Function
Public Function Listado_Combo_Sellantes_de_Clientes(cliente As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT S.ID_SELLANTE AS C1,CONCAT(S.PRODUCTO,' (',C.NOMBRE,')') AS C2 " & _
                   "  FROM SELLANTES S, CLIENTES C " & _
                   " WHERE S.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND S.CLIENTE_ID =" & cliente & _
                   "   AND S.ANULADO = 0 " & _
                   " ORDER BY S.PRODUCTO"
        Set Listado_Combo_Sellantes_de_Clientes = datos_bd(consulta)
End Function


