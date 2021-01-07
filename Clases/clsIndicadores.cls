VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsIndicadores
'***************************************************************
Private ID_INDICADOR As Long
Private NOMBRE As String
Private DESCRIPCION As String
Private FRECUENCIA_ID As Long
Private HOJA_EXCEL As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsIndicadores
'***************************************************************
Public Property Let setID_INDICADOR(ByVal dato As Long)
        ID_INDICADOR = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFRECUENCIA_ID(ByVal dato As Long)
        FRECUENCIA_ID = dato
End Property
Public Property Let setHOJA_EXCEL(ByVal dato As String)
        HOJA_EXCEL = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsIndicadores
'***************************************************************
Public Property Get getID_INDICADOR() As Long
        getID_INDICADOR = ID_INDICADOR
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFRECUENCIA_ID() As Long
        getFRECUENCIA_ID = FRECUENCIA_ID
End Property
Public Property Get getHOJA_EXCEL() As String
        getHOJA_EXCEL = HOJA_EXCEL
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsIndicadores
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores WHERE ID_INDICADOR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_INDICADOR = rs("ID_INDICADOR")
                NOMBRE = rs("NOMBRE")
                DESCRIPCION = rs("DESCRIPCION")
                FRECUENCIA_ID = rs("FRECUENCIA_ID")
                HOJA_EXCEL = rs("HOJA_EXCEL")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsIndicadores)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_INDICADOR) FROM Indicadores"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_INDICADOR = 1
        Else
                ID_INDICADOR = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO Indicadores " & _
                   "  VALUES (" & _
                        ID_INDICADOR & "," & "'" & NOMBRE & "'" & "," & "'" & DESCRIPCION & "'" & "," & _
                        FRECUENCIA_ID & "," & "'" & HOJA_EXCEL & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_INDICADOR
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsIndicadores)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE Indicadores SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FRECUENCIA_ID = " & FRECUENCIA_ID & "," & _
                        " HOJA_EXCEL = '" & HOJA_EXCEL & "'" & _
                " WHERE ID_INDICADOR = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsIndicadores)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM Indicadores " & _
                "    WHERE ID_INDICADOR = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsIndicadores)", vbCritical, Err.Description
End Function
Public Function Listado() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores ORDER BY ID_INDICADOR"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_codigo(ID As Long) As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT i.nombre,f.nombre,i.hoja_excel,i.id_indicador,f.id_frecuencia " & _
                   "  FROM Indicadores i, indicadores_frecuencias f " & _
                   " WHERE i.FRECUENCIA_ID = f.id_frecuencia " & _
                   "   AND i.id_indicador = " & ID
        Set Listado_por_codigo = datos_bd(consulta)
End Function
Public Function lista() As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT i.nombre,f.nombre,i.hoja_excel,i.id_indicador,f.id_frecuencia " & _
                   "  FROM Indicadores i, indicadores_frecuencias f " & _
                   " WHERE i.FRECUENCIA_ID = f.id_frecuencia " & _
                   " ORDER BY i.nombre"
        Set lista = datos_bd(consulta)
End Function

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
        .informe = "General\rptListadoIndicadores"
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

Public Function generar(ID As Long) As adodb.Recordset
        Dim consulta As String
        consulta = "SELECT c.nombre,c.funcion_id,c.datos,d.posicion_excel " & _
                   "  FROM Indicadores i, indicadores_def d,indicadores_campos c " & _
                   " WHERE i.ID_INDICADOR =d.INDICADOR_ID " & _
                   "   AND d.CAMPO_ID = c.ID_CAMPO " & _
                   "   AND i.ID_INDICADOR=" & ID & _
                   " ORDER BY d.orden"
        Set generar = datos_bd(consulta)
End Function
