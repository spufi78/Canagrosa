VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPinturas_radmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPINTURAS_RADMIN
'***************************************************************
Private ID_RADMIN As Long
Private CLIENTE_ID As Long
Private OFERTA_ID As Long
Private PEDIDO_ID As Long
Private FECHA_RECEPCION As String
Private ESTADOS_LATAS As String
Private MARCADO_LATAS As String
Private ESTADO_ID As Integer
Private USUARIO_ID As Integer
Private FS As String
'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSPINTURAS_RADMIN
'***************************************************************
Public Enum LISTA_CLSPINTURAS_RADMIN_COLS
        ENUM_ID_RADMIN = 1
        ENUM_CLIENTE_ID = 2
        ENUM_OFERTA_ID = 3
        ENUM_PEDIDO_ID = 4
        ENUM_FECHA_RECEPCION = 5
        ENUM_ESTADOS_LATAS = 6
        ENUM_MARCADO_LATAS = 7
        ENUM_ESTADO_ID = 8
        ENUM_USUARIO_ID = 9
        ENUM_FS = 10
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPINTURAS_RADMIN
'***************************************************************
Public Property Let setID_RADMIN(ByVal dato As Long)
        ID_RADMIN = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setOFERTA_ID(ByVal dato As Long)
        OFERTA_ID = dato
End Property
Public Property Let setPEDIDO_ID(ByVal dato As Long)
        PEDIDO_ID = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
        FECHA_RECEPCION = dato
End Property
Public Property Let setESTADOS_LATAS(ByVal dato As String)
        ESTADOS_LATAS = dato
End Property
Public Property Let setMARCADO_LATAS(ByVal dato As String)
        MARCADO_LATAS = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Integer)
        ESTADO_ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Integer)
        USUARIO_ID = dato
End Property
Public Property Let setFS(ByVal dato As String)
        FS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPINTURAS_RADMIN
'***************************************************************
Public Property Get getID_RADMIN() As Long
        getID_RADMIN = ID_RADMIN
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getOFERTA_ID() As Long
        getOFERTA_ID = OFERTA_ID
End Property
Public Property Get getPEDIDO_ID() As Long
        getPEDIDO_ID = PEDIDO_ID
End Property
Public Property Get getFECHA_RECEPCION() As String
        getFECHA_RECEPCION = FECHA_RECEPCION
End Property
Public Property Get getESTADOS_LATAS() As String
        getESTADOS_LATAS = ESTADOS_LATAS
End Property
Public Property Get getMARCADO_LATAS() As String
        getMARCADO_LATAS = MARCADO_LATAS
End Property
Public Property Get getESTADO_ID() As Integer
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getUSUARIO_ID() As Integer
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFS() As String
        getFS = FS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPINTURAS_RADMIN
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PINTURAS_RADMIN WHERE ID_RADMIN = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_RADMIN = rs("ID_RADMIN")
                CLIENTE_ID = rs("CLIENTE_ID")
                OFERTA_ID = rs("OFERTA_ID")
                PEDIDO_ID = rs("PEDIDO_ID")
                FECHA_RECEPCION = rs("FECHA_RECEPCION")
                ESTADOS_LATAS = rs("ESTADOS_LATAS")
                MARCADO_LATAS = rs("MARCADO_LATAS")
                ESTADO_ID = rs("ESTADO_ID")
                USUARIO_ID = rs("USUARIO_ID")
                FS = rs("FS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPinturas_radmin)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_RADMIN) FROM PINTURAS_RADMIN"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_RADMIN = 1
        Else
                ID_RADMIN = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PINTURAS_RADMIN " & _
                   "  VALUES (" & _
                        ID_RADMIN & "," & CLIENTE_ID & "," & OFERTA_ID & "," & _
                        PEDIDO_ID & "," & "'" & FECHA_RECEPCION & "'" & "," & _
                        "'" & ESTADOS_LATAS & "'" & "," & "'" & MARCADO_LATAS & "'," & ESTADO_ID & "," & USUARIO_ID & ",CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Insertar = ID_RADMIN
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPinturas_radmin)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PINTURAS_RADMIN SET " & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " OFERTA_ID = " & OFERTA_ID & "," & _
                        " PEDIDO_ID = " & PEDIDO_ID & "," & _
                        " FECHA_RECEPCION = '" & FECHA_RECEPCION & "'," & _
                        " ESTADOS_LATAS = '" & ESTADOS_LATAS & "'," & _
                        " MARCADO_LATAS = '" & MARCADO_LATAS & "'," & _
                        " ESTADO_ID = " & ESTADO_ID & _
                " WHERE ID_RADMIN = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPinturas_radmin)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PINTURAS_RADMIN " & _
                "    WHERE ID_RADMIN = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPinturas_radmin)", vbCritical, Err.Description
End Function
Public Function Listado(fdesde As String, fhasta As String, cCliente As Long, cEstado As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If cCliente <> 0 Then
            filtro = filtro & " AND A.CLIENTE_ID = " & cCliente
        End If
        If cEstado <> 0 Then
            filtro = filtro & " AND A.ESTADO_ID = " & cEstado
        End If
            
        consulta = "SELECT A.ID_RADMIN, B.NOMBRE,A.FECHA_RECEPCION, D.DESCRIPCION,E.USUARIO,A.FS " & _
                   "  FROM PINTURAS_RADMIN A, CLIENTES B, decodificadora D, usuarios E " & _
                   " WHERE FECHA_RECEPCION >= '" & Format(fdesde, "yyyy-mm-dd") & " ' AND FECHA_RECEPCION <= '" & Format(fhasta, "yyyy-mm-dd") & "' " & _
                   "   AND A.CLIENTE_ID = B.ID_CLIENTE " & _
                   "   AND A.ESTADO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.PINTURAS_ESTADOS & _
                   "   AND A.USUARIO_ID = E.ID_EMPLEADO " & _
                   " ORDER BY A.ID_RADMIN DESC"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPINTURAS_RADMIN
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function ListadoID(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_RADMIN, B.NOMBRE,A.FECHA_RECEPCION, D.DESCRIPCION,E.USUARIO,A.FS " & _
                   "  FROM PINTURAS_RADMIN A, CLIENTES B, decodificadora D, usuarios E " & _
                   " WHERE A.CLIENTE_ID = B.ID_CLIENTE " & _
                   "   AND A.ESTADO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.PINTURAS_ESTADOS & _
                   "   AND A.USUARIO_ID = E.ID_EMPLEADO " & _
                   "   AND A.ID_RADMIN = " & ID & _
                   " ORDER BY A.ID_RADMIN DESC"
        Set ListadoID = datos_bd(consulta)
End Function
