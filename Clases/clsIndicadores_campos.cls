VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicadores_campos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsIndicadores_campos
'***************************************************************
Private ID_CAMPO As Long
Private NOMBRE As String
Private DESCRIPCION As String
Private FUNCION_ID As Integer
Private datos As String
Public Sub Imprimir_Listado()

    Dim objRpt As New frmReport
    Dim arrNom() As String
    Dim arrVal() As String
    
    ReDim arrNom(1)
    ReDim arrVal(1)

    arrNom(1) = "impresor"
    arrVal(1) = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        
    With objRpt
        .iniciar
        .informe = "General\rptListadoCamposIndicadores"
        '.CRITERIO = ""
        .imprimir = False
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal

        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objRpt
    Set objRpt = Nothing

End Sub

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsIndicadores_campos
'***************************************************************
Public Property Let setID_CAMPO(ByVal dato As Long)
        ID_CAMPO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFUNCION_ID(ByVal dato As Integer)
        FUNCION_ID = dato
End Property
Public Property Let setDATOS(ByVal dato As String)
        datos = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsIndicadores_campos
'***************************************************************
Public Property Get getID_CAMPO() As Long
        getID_CAMPO = ID_CAMPO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFUNCION_ID() As Integer
        getFUNCION_ID = FUNCION_ID
End Property
Public Property Get getDATOS() As String
        getDATOS = datos
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsIndicadores_campos
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores_campos WHERE ID_CAMPO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_CAMPO = rs("ID_CAMPO")
                NOMBRE = rs("NOMBRE")
                DESCRIPCION = rs("DESCRIPCION")
                FUNCION_ID = rs("FUNCION_ID")
                datos = rs("DATOS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsIndicadores_campos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_CAMPO) FROM Indicadores_campos"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_CAMPO = 1
        Else
                ID_CAMPO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO Indicadores_campos " & _
                   "  VALUES (" & _
                        ID_CAMPO & "," & "'" & NOMBRE & "'" & "," & "'" & DESCRIPCION & "'" & "," & _
                        FUNCION_ID & "," & "'" & datos & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_CAMPO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsIndicadores_campos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE Indicadores_campos SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FUNCION_ID = " & FUNCION_ID & "," & _
                        " DATOS = '" & datos & "'" & _
                   " WHERE ID_CAMPO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsIndicadores_campos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM Indicadores_campos " & _
                "    WHERE ID_CAMPO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsIndicadores_campos)", vbCritical, Err.Description
End Function
Public Function Listado() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores_campos ORDER BY ID_CAMPO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_codigo(ID As Long) As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT i.nombre,f.nombre,i.datos,i.id_campo " & _
                   "  FROM Indicadores_campos i, indicadores_funciones f " & _
                   " WHERE i.funcion_id = f.id_funcion " & _
                   "   AND i.id_campo = " & ID
        Set Listado_por_codigo = datos_bd(consulta)
End Function
Public Function lista() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT i.nombre,f.nombre,i.datos,i.id_campo " & _
                   "  FROM Indicadores_campos i, indicadores_funciones f " & _
                   " WHERE i.funcion_id = f.id_funcion " & _
                   " ORDER BY i.nombre"
        Set lista = datos_bd(consulta)
End Function
Public Function Listado_Combo() As adodb.Recordset
    Dim consulta As String
    consulta = "SELECT ID_CAMPO,NOMBRE FROM Indicadores_campos order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
