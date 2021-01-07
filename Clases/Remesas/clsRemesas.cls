VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSREMESAS
'***************************************************************
Private ID_REMESA As Long
Private fecha As String
Private NUMERO As Integer
Private ANNO As Integer
Private DESCRIPCION As String
Private Documentos As Long
Private IMPORTE As String
Private TIPO_ID As Long
Private ESTADO_ID As Integer
Private BANCO_ID As Long
Private USUARIO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSREMESAS
'***************************************************************
Public Property Let setID_REMESA(ByVal dato As Long)
        ID_REMESA = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setNUMERO(ByVal dato As Integer)
        NUMERO = dato
End Property
Public Property Let setANNO(ByVal dato As Integer)
        ANNO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setDOCUMENTOS(ByVal dato As Long)
        Documentos = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Integer)
        ESTADO_ID = dato
End Property
Public Property Let setBANCO_ID(ByVal dato As Long)
        BANCO_ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSREMESAS
'***************************************************************
Public Property Get getID_REMESA() As Long
        getID_REMESA = ID_REMESA
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getNUMERO() As Integer
        getNUMERO = NUMERO
End Property
Public Property Get getANNO() As Integer
        getANNO = ANNO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getDOCUMENTOS() As Long
        getDOCUMENTOS = Documentos
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getESTADO_ID() As Integer
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getBANCO_ID() As Long
        getBANCO_ID = BANCO_ID
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSREMESAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REMESAS WHERE ID_REMESA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_REMESA = rs("ID_REMESA")
                fecha = rs("FECHA")
                NUMERO = rs("NUMERO")
                ANNO = rs("ANNO")
                DESCRIPCION = rs("DESCRIPCION")
                Documentos = rs("DOCUMENTOS")
                IMPORTE = rs("IMPORTE")
                TIPO_ID = rs("TIPO_ID")
                ESTADO_ID = rs("ESTADO_ID")
                BANCO_ID = rs("BANCO_ID")
                USUARIO_ID = rs("USUARIO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRemesas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_REMESA) FROM REMESAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_REMESA = 1
        Else
                ID_REMESA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function CrearNumero()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(NUMERO) FROM REMESAS WHERE ANNO = YEAR(CURRENT_DATE)"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                NUMERO = 1
        Else
                NUMERO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        Me.CrearNumero
        consulta = "INSERT INTO REMESAS " & _
                   "  VALUES (" & _
                        ID_REMESA & ",CURRENT_TIMESTAMP," & NUMERO & ",YEAR(CURRENT_DATE)," & "'" & DESCRIPCION & "'" & "," & _
                        Documentos & ",'" & IMPORTE & "'," & TIPO_ID & "," & ESTADO_ID & "," & BANCO_ID & "," & USUARIO.getID_EMPLEADO & _
                ")"
        execute_bd consulta
        Insertar = ID_REMESA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRemesas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE REMESAS SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " DOCUMENTOS = " & Documentos & "," & _
                        " IMPORTE = '" & IMPORTE & "'," & _
                        " TIPO_ID = " & TIPO_ID & "," & _
                        " ESTADO_ID = " & ESTADO_ID & "," & _
                        " BANCO_ID = " & BANCO_ID & "" & _
                " WHERE ID_REMESA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRemesas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM REMESAS_DOCUMENTOS WHERE REMESA_ID = " & ID
        execute_bd consulta
        consulta = "DELETE FROM REMESAS WHERE ID_REMESA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRemesas)", vbCritical, Err.Description
End Function
Public Function Listado(idRemesa As String, USUARIO As String, fdesde As String, fhasta As String, BANCO_ID As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If idRemesa <> "" Then
            If IsNumeric(idRemesa) Then
                filtro = filtro & " AND A.ID_REMESA = " & idRemesa
            End If
        End If
        If USUARIO <> "" Then
            filtro = filtro & " AND B.usuario LIKE '%" & USUARIO & "%'"
        End If
        If fdesde <> "" Then
            filtro = filtro & " AND A.FECHA >= '" & fdesde & "'"
        End If
        If fhasta <> "" Then
            filtro = filtro & " AND A.FECHA <= '" & fhasta & "'"
        End If
        If BANCO_ID <> 0 Then
            filtro = filtro & " AND A.BANCO_ID = " & BANCO_ID
        End If
        consulta = "SELECT A.ID_REMESA,CONCAT(A.NUMERO,'/',A.ANNO),A.FECHA,C.DESCRIPCION, A.DESCRIPCION,A.DOCUMENTOS,A.IMPORTE,B.USUARIO,D.DESCRIPCION " & _
                   "  FROM REMESAS A, usuarios B, BANCOS C, decodificadora D " & _
                   " WHERE A.USUARIO_ID = B.ID_EMPLEADO " & _
                   "   AND A.BANCO_ID = C.ID_BANCO " & _
                   "   AND A.ESTADO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_REMESAS_ESTADOS & _
                   filtro & _
                   " ORDER BY ID_REMESA DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoIds(idRemesa As Long) As String
        Dim consulta As String
   On Error GoTo ListadoIds_Error

'        consulta = "select distinct factura_id from remesas_documentos where remesa_id = " & idRemesa
        consulta = "select group_concat(factura_id) from remesas_documentos where remesa_id = " & idRemesa
        Dim rs As ADODB.Recordset
        Dim lista As String
        lista = ""
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
'            Do
'                If lista <> "" Then
'                    lista = lista & ","
'                End If
'                lista = lista & rs(0)
'                rs.MoveNext
'            Loop Until rs.EOF
            lista = Trim(rs(0))
        End If
        ListadoIds = lista

   On Error GoTo 0
   Exit Function

ListadoIds_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListadoIds of Módulo de clase clsRemesas"
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSREMESAS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_REMESA,FECHA FROM REMESAS ORDER BY FECHA"
        Set Listado_Combo = datos_bd(consulta)
End Function
