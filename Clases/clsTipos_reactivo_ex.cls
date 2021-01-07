VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_reactivo_ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_TIPO_REACTIVO_EX As Integer
Private nombre As String
Private CERTIFICADO As String
Private ALMACENAJE As String
Private SEGURIDAD As String
Private FDS As String
Private anulado As Integer
Private PROBETA As Integer
Public Property Let setID_TIPO_REACTIVO_EX(ByVal dato As Integer)
    ID_TIPO_REACTIVO_EX = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setCERTIFICADO(ByVal dato As String)
    CERTIFICADO = Trim(dato)
End Property
Public Property Let setALMACENAJE(ByVal dato As String)
    ALMACENAJE = Trim(dato)
End Property
Public Property Let setSEGURIDAD(ByVal dato As String)
    SEGURIDAD = Trim(dato)
End Property
Public Property Let setFDS(ByVal dato As String)
    FDS = Trim(dato)
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
Public Property Let setPROBETA(ByVal dato As Integer)
    PROBETA = dato
End Property
Public Property Get getID_TIPO_REACTIVO_EX() As Integer
    getID_TIPO_REACTIVO_EX = ID_TIPO_REACTIVO_EX
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getCERTIFICADO() As String
    getCERTIFICADO = CERTIFICADO
End Property
Public Property Get getALMACENAJE() As String
    getALMACENAJE = ALMACENAJE
End Property
Public Property Get getSEGURIDAD() As String
    getSEGURIDAD = SEGURIDAD
End Property
Public Property Get getFDS() As String
    getFDS = FDS
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
Public Property Get getPROBETA() As Integer
    getPROBETA = PROBETA
End Property
Public Function Listado(DESCRIPCION As String, FDS As String, ANULADOS As Boolean, PROBETA As Boolean) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_REACTIVO_EX " & _
               " WHERE NOMBRE LIKE '%" & DESCRIPCION & "%'" & _
               "   AND FDS LIKE '%" & FDS & "%'" & _
               IIf(ANULADOS, " AND ANULADO <> 0 ", " AND ANULADO = 0 ") & _
               IIf(PROBETA, " AND PROBETA = 1 ", "") & _
               " ORDER BY NOMBRE"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Filtrado(texto As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT NOMBRE, ID_TIPO_REACTIVO_EX FROM TIPOS_REACTIVO_EX WHERE NOMBRE LIKE '%" & texto & "%' order by NOMBRE"
    Set Listado_Filtrado = datos_bd(consulta)
End Function
Public Function CARGAR(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_REACTIVO_EX WHERE ID_TIPO_REACTIVO_EX=" & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID_TIPO_REACTIVO_EX = rs("ID_TIPO_REACTIVO_EX")
        nombre = rs("NOMBRE")
        CERTIFICADO = rs("CERTIFICADO")
        ALMACENAJE = rs("ALMACENAJE")
        SEGURIDAD = rs("SEGURIDAD")
        FDS = rs("FDS")
        anulado = rs("ANULADO")
        PROBETA = rs("PROBETA")
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
    MsgBox "Error al cargar el reactivo.", vbCritical, Err.Description
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_REACTIVO_EX) FROM TIPOS_REACTIVO_EX"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_REACTIVO_EX = 1
    Else
        ID_TIPO_REACTIVO_EX = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
'    If ID_TIPO_REACTIVO_EX = 0 Then
    If duplicado = False Then
        Me.CrearID
'    End If
        consulta = "Insert into TIPOS_REACTIVO_EX " & _
                   " values(" & _
                   ID_TIPO_REACTIVO_EX & ",'" & nombre & "','" & _
                   CERTIFICADO & "','" & _
                   ALMACENAJE & "','" & _
                   SEGURIDAD & "','" & _
                   FDS & "'," & _
                   anulado & "," & _
                   PROBETA & ")"
        execute_bd consulta
        Insertar = ID_TIPO_REACTIVO_EX
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Reactivo (Insertar)", vbCritical, Err.Description
End Function
Public Sub Imprimir_Listado(ByVal prm_sustancia As String, ByVal prm_FDS As String, ByVal prm_anulados As Boolean)

    Dim objRpt As New frmReport
    Dim arrNom() As String
    Dim arrVal() As String
    Dim strCriterio As String
                
    If prm_anulados Then strCriterio = strCriterio & IIf(Trim(strCriterio) <> "", " AND", "") & " {tipos_reactivo_ex.ANULADO} = 0"
    If Trim(prm_sustancia) <> "" Then strCriterio = strCriterio & IIf(Trim(strCriterio) <> "", " AND", "") & " {tipos_reactivo_ex.NOMBRE} like " & Chr(34) & "*" & prm_sustancia & "*" & Chr(34)
    If Trim(prm_FDS) <> "" Then strCriterio = strCriterio & IIf(Trim(strCriterio) <> "", " AND", "") & " {tipos_reactivo_ex.FDS} like " & Chr(34) & "*" & prm_FDS & "*" & Chr(34)
                
    ReDim arrNom(1)
    ReDim arrVal(1)

    arrNom(1) = "impresor"
    arrVal(1) = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        
    With objRpt
        .iniciar
        .informe = "General\rptListadoReactivosExternos"
        .criterio = strCriterio
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

Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
'    Dim rs As New ADODB.RecordSet
    Dim consulta As String
'    Set rs = datos_bd("SELECT ID_TIPO_BOTE_EX FROM TIPOS_BOTE_EX WHERE TIPO_REACTIVO_EX_ID = " & ID)
'    If rs.RecordCount > 0 Then
'        MsgBox "No puede eliminarse el tipo de reactivo. Se usa en el bote con ID = " & rs(0)
'        Eliminar = False
'    Else
'        consulta = "DELETE FROM TIPOS_REACTIVO_EX WHERE ID_TIPO_REACTIVO_EX=" & ID
        consulta = "UPDATE TIPOS_REACTIVO_EX SET ANULADO = " & USUARIO.getID_EMPLEADO & " WHERE ID_TIPO_REACTIVO_EX=" & ID
        execute_bd consulta
        Eliminar = True
'    End If
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el Reactivo (Eliminar)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE TIPOS_REACTIVO_EX SET " & _
               " NOMBRE = '" & nombre & "'," & _
               " CERTIFICADO = '" & CERTIFICADO & "'," & _
               " ALMACENAJE = '" & ALMACENAJE & "'," & _
               " FDS = '" & FDS & "'," & _
               " SEGURIDAD = '" & SEGURIDAD & "'," & _
               " PROBETA = " & PROBETA & _
               " WHERE ID_TIPO_REACTIVO_EX = " & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar el Reactivo (Modificar)", vbCritical, Err.Description
End Function
Public Function duplicado() As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_REACTIVO_EX FROM TIPOS_REACTIVO_EX WHERE NOMBRE = '" & nombre & "'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        duplicado = False
    Else
        ID_TIPO_REACTIVO_EX = rs(0)
        duplicado = True
    End If
    Exit Function
fallo:
    duplicado = False
End Function

Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT ID_TIPO_REACTIVO_EX, NOMBRE " & _
               "  FROM TIPOS_REACTIVO_EX " & _
               " ORDER BY NOMBRE ASC"
    Set Listado_Combo = datos_bd(consulta)
End Function

Public Function Listado_Pictogramas(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.NOMBRE,A.ID_PICTOGRAMA,A.IMAGEN " & _
               "  FROM PICTOGRAMAS AS A, REACTIVOS_EX_PICTOGRAMAS AS B " & _
               " WHERE A.ID_PICTOGRAMA = B.PICTOGRAMA_ID " & _
               "   AND B.REACTIVO_EX_ID = " & ID
    Set Listado_Pictogramas = datos_bd(consulta)
End Function
Public Function Listado_Frases(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.CODIGO, A.FRASE,A.ID_FRASE " & _
               "  FROM FRASESRYS AS A, REACTIVOS_EX_FRASESRYS AS B " & _
               " WHERE A.ID_FRASE = B.FRASE_ID " & _
               "   AND B.REACTIVO_EX_ID = " & ID
    Set Listado_Frases = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "TIPOS_REACTIVO_EX"
        .setDESCRIPCION = "Sustancias / Materiales"
        .setPK = "ID_TIPO_REACTIVO_EX"
        .setCAMPO = "NOMBRE"
        .setFILTRO = filtro
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
