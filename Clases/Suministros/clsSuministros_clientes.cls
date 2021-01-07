VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuministros_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSUMINISTROS_CLIENTES
'***************************************************************
Private TIPO_SUMINISTRO_ID As Long
Private ORDEN As Long
Private CLIENTE_ID As Long
Private CAPACIDAD_ID As Long
Private NUMERO_BOTES As Long
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSUMINISTROS_CLIENTES
'***************************************************************
Public Property Let setTIPO_SUMINISTRO_ID(ByVal dato As Long)
        TIPO_SUMINISTRO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setCAPACIDAD_ID(ByVal dato As Long)
        CAPACIDAD_ID = dato
End Property
Public Property Let setNUMERO_BOTES(ByVal dato As Long)
        NUMERO_BOTES = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSUMINISTROS_CLIENTES
'***************************************************************
Public Property Get getTIPO_SUMINISTRO_ID() As Long
        getTIPO_SUMINISTRO_ID = TIPO_SUMINISTRO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getCAPACIDAD_ID() As Long
        getCAPACIDAD_ID = CAPACIDAD_ID
End Property
Public Property Get getNUMERO_BOTES() As Long
        getNUMERO_BOTES = NUMERO_BOTES
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSUMINISTROS_CLIENTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SUMINISTROS_CLIENTES WHERE TIPO_SUMINISTRO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TIPO_SUMINISTRO_ID = rs("TIPO_SUMINISTRO_ID")
                ORDEN = rs("ORDEN")
                CLIENTE_ID = rs("CLIENTE_ID")
                CAPACIDAD_ID = rs("CAPACIDAD_ID")
                NUMERO_BOTES = rs("NUMERO_BOTES")
                PRECIO = rs("PRECIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSuministros_clientes)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO SUMINISTROS_CLIENTES " & _
                   "  VALUES (" & _
                        TIPO_SUMINISTRO_ID & "," & ORDEN & "," & CLIENTE_ID & "," & CAPACIDAD_ID & "," & _
                        NUMERO_BOTES & ",'" & PRECIO & "'" & _
                ")"
        execute_bd consulta
        Insertar = TIPO_SUMINISTRO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSuministros_clientes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SUMINISTROS_CLIENTES SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " CAPACIDAD_ID = " & CAPACIDAD_ID & "," & _
                        " NUMERO_BOTES = " & NUMERO_BOTES & "," & _
                        " PRECIO = '" & PRECIO & "'" & _
                " WHERE TIPO_SUMINISTRO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSuministros_clientes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SUMINISTROS_CLIENTES " & _
                "    WHERE TIPO_SUMINISTRO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSuministros_clientes)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.NOMBRE,C.DESCRIPCION,A.NUMERO_BOTES,A.CLIENTE_ID,A.CAPACIDAD_ID,A.PRECIO " & _
                   "  FROM SUMINISTROS_CLIENTES A,CLIENTES B, ALODINE_CAPACIDAD C " & _
                   " WHERE A.CLIENTE_ID = B.ID_CLIENTE " & _
                   "   AND A.CAPACIDAD_ID = C.ID_CAPACIDAD " & _
                   "   AND A.TIPO_SUMINISTRO_ID = " & ID & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function
