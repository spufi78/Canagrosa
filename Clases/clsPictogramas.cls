VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPictogramas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsPictogramas
'***************************************************************
Private ID_PICTOGRAMA As Long
Private NOMBRE As String
Private imagen As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsPictogramas
'***************************************************************
Public Property Let setID_PICTOGRAMA(ByVal dato As Long)
        ID_PICTOGRAMA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setIMAGEN(ByVal dato As String)
        imagen = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsPictogramas
'***************************************************************
Public Property Get getID_PICTOGRAMA() As Long
        getID_PICTOGRAMA = ID_PICTOGRAMA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getIMAGEN() As String
        getIMAGEN = imagen
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsPictogramas
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Pictogramas WHERE ID_PICTOGRAMA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PICTOGRAMA = rs("ID_PICTOGRAMA")
                NOMBRE = rs("NOMBRE")
                imagen = rs("IMAGEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPictogramas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_PICTOGRAMA) FROM Pictogramas"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PICTOGRAMA = 1
        Else
                ID_PICTOGRAMA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO Pictogramas " & _
                   "  VALUES (" & _
                        ID_PICTOGRAMA & "," & "'" & NOMBRE & "'" & "," & "'" & imagen & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_PICTOGRAMA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPictogramas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE Pictogramas SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " IMAGEN = '" & imagen & "'" & _
                " WHERE ID_PICTOGRAMA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPictogramas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM Pictogramas " & _
                "    WHERE ID_PICTOGRAMA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPictogramas)", vbCritical, Err.Description
End Function
Public Function Listado() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Pictogramas ORDER BY ID_PICTOGRAMA"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As adodb.Recordset
    Dim consulta As String
    consulta = "SELECT ID_PICTOGRAMA,NOMBRE FROM PICTOGRAMAS order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
