VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrasesrys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsFrasesrys
'***************************************************************
Private ID_FRASE As Long
Private CODIGO As String
Private FRASE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsFrasesrys
'***************************************************************
Public Property Let setID_FRASE(ByVal dato As Long)
        ID_FRASE = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setFRASE(ByVal dato As String)
        FRASE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsFrasesrys
'***************************************************************
Public Property Get getID_FRASE() As Long
        getID_FRASE = ID_FRASE
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getFRASE() As String
        getFRASE = FRASE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsFrasesrys
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Frasesrys WHERE ID_FRASE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FRASE = rs("ID_FRASE")
                CODIGO = rs("CODIGO")
                FRASE = rs("FRASE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFrasesrys)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_FRASE) FROM Frasesrys"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FRASE = 1
        Else
                ID_FRASE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO Frasesrys " & _
                   "  VALUES (" & _
                        ID_FRASE & "," & "'" & CODIGO & "'" & "," & "'" & FRASE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_FRASE
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFrasesrys)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE Frasesrys SET " & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " FRASE = '" & FRASE & "'" & _
                " WHERE ID_FRASE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFrasesrys)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM Frasesrys " & _
                "    WHERE ID_FRASE = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFrasesrys)", vbCritical, Err.Description
End Function
Public Function Listado() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Frasesrys ORDER BY ID_FRASE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As adodb.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FRASE,concat('(',CODIGO,') ',FRASE) AS NOMBRE FROM frasesrys order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
