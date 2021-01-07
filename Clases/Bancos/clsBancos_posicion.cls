VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBancos_posicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSBANCOS_POSICION
'***************************************************************
Private ID As Long
Private BANCO_ID As Long
Private fecha As String
Private DESCRIPCION As String
Private IMPORTE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSBANCOS_POSICION
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setBANCO_ID(ByVal dato As Long)
        BANCO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSBANCOS_POSICION
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getBANCO_ID() As Long
        getBANCO_ID = BANCO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSBANCOS_POSICION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM BANCOS_POSICION WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                BANCO_ID = rs("BANCO_ID")
                fecha = rs("FECHA")
                DESCRIPCION = rs("DESCRIPCION")
                IMPORTE = rs("IMPORTE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsBancos_posicion)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM BANCOS_POSICION"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO BANCOS_POSICION " & _
                   "  VALUES (" & _
                        ID & "," & BANCO_ID & "," & "'" & fecha & "'" & "," & _
                        "'" & DESCRIPCION & "'" & ",'" & IMPORTE & "'" & _
                ")"
        execute_bd consulta
        actualizarSaldoBanco BANCO_ID
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsBancos_posicion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE BANCOS_POSICION SET " & _
                        " BANCO_ID = " & BANCO_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " IMPORTE = '" & IMPORTE & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        actualizarSaldoBanco BANCO_ID
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsBancos_posicion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim oBP As New clsBancos_posicion
        oBP.Carga ID
        consulta = "DELETE FROM BANCOS_POSICION " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        actualizarSaldoBanco oBP.getBANCO_ID
        Set oBP = Nothing
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsBancos_posicion)", vbCritical, Err.Description
End Function
Public Function Listado(ID_BANCO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID,A.FECHA,A.DESCRIPCION,A.IMPORTE " & _
                   "  FROM BANCOS_POSICION A " & _
                   " WHERE BANCO_ID = " & ID_BANCO & _
                   " ORDER BY A.FECHA DESC "
        Set Listado = datos_bd(consulta)
End Function
Private Function actualizarSaldoBanco(ID_BANCO As Long) As Boolean
        Dim rs As ADODB.Recordset
        Dim consulta As String
        Dim IMPORTE As Currency
   On Error GoTo actualizarSaldoBanco_Error

        IMPORTE = 0
        consulta = "SELECT IMPORTE FROM BANCOS_POSICION WHERE BANCO_ID=" & ID_BANCO & " ORDER BY FECHA DESC LIMIT 1"
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            IMPORTE = rs(0)
        End If
        Set rs = Nothing
        Dim oBanco As New clsBancos
        oBanco.actualizarSaldo ID_BANCO, moneda_bd(CStr(IMPORTE))
        Set oBanco = Nothing

   On Error GoTo 0
   Exit Function

actualizarSaldoBanco_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizarSaldoBanco of Módulo de clase clsBancos_posicion"
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSBANCOS_POSICION
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,BANCO_ID FROM BANCOS_POSICION ORDER BY BANCO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
