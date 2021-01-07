VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubfamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsSUBFAMILIAS
'***************************************************************
Private ID_SUBFAMILIA As Integer
Private FAMILIA_ID As Integer
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsSUBFAMILIAS
'***************************************************************
Public Property Let setID_SUBFAMILIA(ByVal dato As Integer)
        ID_SUBFAMILIA = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Integer)
        FAMILIA_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsSUBFAMILIAS
'***************************************************************
Public Property Get getID_SUBFAMILIA() As Integer
        getID_SUBFAMILIA = ID_SUBFAMILIA
End Property
Public Property Get getFAMILIA_ID() As Integer
        getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsSUBFAMILIAS
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SUBFAMILIAS WHERE ID_SUBFAMILIA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SUBFAMILIA = rs("ID_SUBFAMILIA")
                FAMILIA_ID = rs("FAMILIA_ID")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSUBFAMILIAS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO SUBFAMILIAS " & _
                   "  VALUES (" & _
                        ID_SUBFAMILIA & "," & FAMILIA_ID & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd consulta
        Insertar = FAMILIA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSUBFAMILIAS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SUBFAMILIAS SET " & _
                        " FAMILIA_ID = " & FAMILIA_ID & "," & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_SUBFAMILIA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSUBFAMILIAS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        ' Validar articulo
        Dim rs As ADODB.Recordset
'        Dim oArticulo As New clsArticulo
'        Set rs = oArticulo.Listado("", "", CStr(ID), "", False)
'        If rs.RecordCount > 0 Then
'            MsgBox "Existen artículos asociados a esta subfamilia. No se puede eliminar.", vbInformation, App.Title
'            Exit Function
'        End If
        consulta = "DELETE FROM SUBFAMILIAS " & _
                "    WHERE ID_SUBFAMILIA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSUBFAMILIAS)", vbCritical, Err.Description
End Function
Public Function Listado_por_familia(ID As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SUBFAMILIAS WHERE FAMILIA_ID = " & ID & " ORDER BY NOMBRE"
        Set Listado_por_familia = datos_bd(consulta)
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SUBFAMILIAS ORDER BY NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_SUBFAMILIA) FROM SUBFAMILIAS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_SUBFAMILIA = 1
    Else
        ID_SUBFAMILIA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub

