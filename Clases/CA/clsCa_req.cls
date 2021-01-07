VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_req"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_REQ
'***************************************************************
Private ID_REQ As Long
Private FAMILIA_REQ_ID As Long
Private nombre As String
Private ORDEN As Long
Private anulado As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_REQ
'***************************************************************
Public Property Let setID_REQ(ByVal dato As Long)
        ID_REQ = dato
End Property
Public Property Let setFAMILIA_REQ_ID(ByVal dato As Long)
        FAMILIA_REQ_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setANULADO(ByVal dato As Long)
        anulado = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_REQ
'***************************************************************
Public Property Get getID_REQ() As Long
        getID_REQ = ID_REQ
End Property
Public Property Get getFAMILIA_REQ_ID() As Long
        getFAMILIA_REQ_ID = FAMILIA_REQ_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getANULADO() As Long
        getANULADO = anulado
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_REQ
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_REQ WHERE ID_REQ = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_REQ = rs("ID_REQ")
                FAMILIA_REQ_ID = rs("FAMILIA_REQ_ID")
                nombre = rs("NOMBRE")
                ORDEN = rs("ORDEN")
                anulado = rs("ANULADO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_req)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_REQ) FROM CA_REQ"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_REQ = 1
        Else
                ID_REQ = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function CrearORDEN(idFam As Integer)
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ORDEN) FROM CA_REQ WHERE FAMILIA_REQ_ID = " & idFam
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ORDEN = 1
        Else
                ORDEN = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        Me.CrearORDEN (FAMILIA_REQ_ID)
        Me.setANULADO = 0
        consulta = "INSERT INTO CA_REQ " & _
                   "  VALUES (" & _
                        ID_REQ & "," & FAMILIA_REQ_ID & "," & "'" & nombre & "'" & "," & _
                        ORDEN & "," & anulado & _
                ")"
        execute_bd consulta
        Insertar = ID_REQ
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_req)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_REQ SET " & _
                        " NOMBRE = '" & nombre & "'" & _
                   " WHERE ID_REQ = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_req)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CA_REQ " & _
                "    WHERE ID_REQ = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_req)", vbCritical, Err.Description
End Function
Public Function Listado(ID_FAMILIA As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_REQ, NOMBRE FROM CA_REQ " & _
                   " WHERE FAMILIA_REQ_ID = " & ID_FAMILIA & _
                   "   AND ANULADO = 0 " & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSCA_REQ
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_REQ,FAMILIA_REQ_ID FROM CA_REQ ORDER BY FAMILIA_REQ_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function duplicar(ID_ORIGEN As Long, ID_DESTINO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("SELECT * FROM CA_REQ WHERE FAMILIA_REQ_ID = " & CStr(ID_ORIGEN) & " ORDER BY ORDEN ")
        If rs.RecordCount > 0 Then
            Do
                Me.setFAMILIA_REQ_ID = ID_DESTINO
                Me.setNOMBRE = rs("NOMBRE")
                Me.Insertar
                rs.MoveNext
            Loop Until rs.EOF
        End If
        duplicar = True
        Exit Function
fallo:
        duplicar = False
        MsgBox "Error al Modificar (clsCa_req)", vbCritical, Err.Description
End Function
Public Function Ordenar(ID_REQ As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim oCR As New clsCa_req
        oCR.Carga ID_REQ
        Dim rs As ADODB.Recordset
        consulta = "SELECT * FROM CA_REQ " & _
                   " WHERE FAMILIA_REQ_ID = " & oCR.getFAMILIA_REQ_ID & _
                   IIf(ORDEN = 1, "   AND ORDEN > " & oCR.getORDEN, "   AND ORDEN < " & oCR.getORDEN) & _
                   " ORDER BY ORDEN " & _
                   IIf(ORDEN = 0, " DESC ", "") & _
                   " LIMIT 1 "
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            execute_bd "UPDATE CA_REQ SET ORDEN = " & oCR.getORDEN & " WHERE ID_REQ = " & rs("ID_REQ")
            execute_bd "UPDATE CA_REQ SET ORDEN = " & rs("ORDEN") & " WHERE ID_REQ = " & ID_REQ
        End If
        Ordenar = True
        Exit Function
fallo:
        Ordenar = False
        MsgBox "Error al Ordenar (clsCa_req)", vbCritical, Err.Description
End Function


