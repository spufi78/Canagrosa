VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHenkel_cfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSHENKEL_CFM
'***************************************************************
Private ID As Long
Private P_INICIO As Long
Private P_FIN As Long
Private PRECIO As String
Private tasa As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSHENKEL_CFM
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setP_INICIO(ByVal dato As Long)
        P_INICIO = dato
End Property
Public Property Let setP_FIN(ByVal dato As Long)
        P_FIN = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setTASA(ByVal dato As String)
        tasa = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSHENKEL_CFM
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getP_INICIO() As Long
        getP_INICIO = P_INICIO
End Property
Public Property Get getP_FIN() As Long
        getP_FIN = P_FIN
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getTASA() As String
        getTASA = tasa
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSHENKEL_CFM
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM HENKEL_CFM WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                P_INICIO = rs("P_INICIO")
                P_FIN = rs("P_FIN")
                PRECIO = rs("PRECIO")
                tasa = rs("TASA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsHenkel_cfm)", vbCritical, Err.Description
End Function
Public Function CargaRango(CANTIDAD As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM HENKEL_CFM WHERE P_INICIO <= " & CANTIDAD & " and P_FIN >= " & CANTIDAD
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                CargaRango = False
        Else
                ID = rs("ID")
                P_INICIO = rs("P_INICIO")
                P_FIN = rs("P_FIN")
                PRECIO = rs("PRECIO")
                tasa = rs("TASA")
                CargaRango = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaRango = False
        MsgBox "Error en CargaRango los datos(clsHenkel_cfm)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM HENKEL_CFM"
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
        consulta = "INSERT INTO HENKEL_CFM " & _
                   "  VALUES (" & _
                        ID & "," & P_INICIO & "," & P_FIN & ",'" & _
                        PRECIO & "','" & tasa & _
                "')"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsHenkel_cfm)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE HENKEL_CFM SET " & _
                        " P_INICIO = " & P_INICIO & "," & _
                        " P_FIN = " & P_FIN & "," & _
                        " PRECIO = '" & PRECIO & "'," & _
                        " TASA = '" & tasa & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsHenkel_cfm)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM HENKEL_CFM " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsHenkel_cfm)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM HENKEL_CFM ORDER BY ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSHENKEL_CFM
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,P_INICIO FROM HENKEL_CFM ORDER BY P_INICIO"
        Set Listado_Combo = datos_bd(consulta)
End Function
