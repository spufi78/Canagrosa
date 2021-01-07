VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_fexternos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE PROVEEDOR    *
'*****************************************************
Private ID_FEXTERNO As Long
Private NOMBRE As String
Private DIRECCION As String
Private COD_POSTAL As Long
Private CIF As String
Private TELEFONO As String
Private FAX As String
Private responsable As String
Private CARGO As String
Private tipo As String
Private EMAIL As String
Private OBSERVACIONES As String
Private WEB As String
Private MUNICIPIO_ID As Long
Private PROVINCIA_ID As Long
Private PAIS_ID As Long
Private ADJUNTOS As String
'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE MUESTRA    *
'*****************************************************
Public Property Let setID_FEXTERNO(ByVal dato As Long)
    ID_FEXTERNO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setDIRECCION(ByVal dato As String)
    DIRECCION = Trim(dato)
End Property

Public Property Let setCOD_POSTAL(ByVal dato As Long)
    COD_POSTAL = dato
End Property
Public Property Let setCIF(ByVal dato As String)
    CIF = Trim(dato)
End Property
Public Property Let setTELEFONO(ByVal dato As String)
    TELEFONO = Trim(dato)
End Property
Public Property Let setFAX(ByVal dato As String)
    FAX = Trim(dato)
End Property
Public Property Let setRESPONSABLE(ByVal dato As String)
    responsable = Trim(dato)
End Property
Public Property Let setCARGO(ByVal dato As String)
    CARGO = Trim(dato)
End Property
Public Property Let setTIPO(ByVal dato As String)
    tipo = Trim(dato)
End Property
Public Property Let setEMAIL(ByVal dato As String)
    EMAIL = Trim(dato)
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setWEB(ByVal dato As String)
    WEB = Trim(dato)
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
    MUNICIPIO_ID = dato
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
    PROVINCIA_ID = dato
End Property
Public Property Let setPAIS_ID(ByVal dato As Long)
    PAIS_ID = dato
End Property
Public Property Let setADJUNTOS(ByVal dato As String)
    ADJUNTOS = dato
End Property
'*****************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE MUESTRA      *
'*****************************************************
Public Property Get getID_FEXTERNO() As Long
    getID_FEXTERNO = ID_FEXTERNO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getDIRECCION() As String
    getDIRECCION = DIRECCION
End Property
Public Property Get getCOD_POSTAL() As Long
    getCOD_POSTAL = COD_POSTAL
End Property
Public Property Get getCIF() As String
    getCIF = CIF
End Property
Public Property Get getTELEFONO() As String
    getTELEFONO = TELEFONO
End Property
Public Property Get getFAX() As String
    getFAX = FAX
End Property
Public Property Get getRESPONSABLE() As String
    getRESPONSABLE = responsable
End Property
Public Property Get getCARGO() As String
    getCARGO = CARGO
End Property
Public Property Get getTIPO() As String
    getTIPO = tipo
End Property
Public Property Get getEMAIL() As String
    getEMAIL = EMAIL
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getWEB() As String
    getWEB = WEB
End Property
Public Property Get getMUNICIPIO_ID() As Long
    getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getPROVINCIA_ID() As Long
    getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getPAIS_ID() As Long
    getPAIS_ID = PAIS_ID
End Property
Public Property Get getADJUNTOS() As String
    getADJUNTOS = ADJUNTOS
End Property
'*****************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE MUESTRA  *
'*****************************************************
Public Function Carga(ID_FEXTERNO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM empleados_fexternos WHERE ID_FEXTERNO=" & ID_FEXTERNO
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            ID_FEXTERNO = rs.Fields("ID_FEXTERNO")
            NOMBRE = rs.Fields("NOMBRE") & ""
            DIRECCION = rs.Fields("DIRECCION") & ""
            If Not IsNull(rs.Fields("COD_POSTAL")) Then
                COD_POSTAL = rs.Fields("COD_POSTAL")
            Else
                COD_POSTAL = 0
            End If
            CIF = rs.Fields("CIF") & ""
            TELEFONO = rs.Fields("TELEFONO") & ""
            FAX = rs.Fields("FAX") & ""
            responsable = rs.Fields("RESPONSABLE") & ""
            CARGO = rs.Fields("CARGO") & ""
            tipo = rs.Fields("TIPO") & ""
            EMAIL = rs.Fields("EMAIL") & ""
            OBSERVACIONES = rs.Fields("OBSERVACIONES") & ""
            WEB = rs.Fields("WEB") & ""
            If IsNull(rs.Fields("MUNICIPIO_ID")) Then
               MUNICIPIO_ID = 0
            Else
               MUNICIPIO_ID = rs.Fields("MUNICIPIO_ID")
            End If
            If IsNull(rs.Fields("PROVINCIA_ID")) Then
               PROVINCIA_ID = 0
            Else
               PROVINCIA_ID = rs.Fields("PROVINCIA_ID")
            End If
            If IsNull(rs.Fields("PAIS_ID")) Then
               PAIS_ID = 0
            Else
               PAIS_ID = rs.Fields("PAIS_ID")
            End If
            ADJUNTOS = rs("ADJUNTOS")
            Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos del formador (Carga)", vbCritical, Err.Description
End Function

Public Function Insertar() As Boolean
    Dim CONSULTA As String
    On Error GoTo fallo:
    If duplicado = False Then
        Me.CrearID_FEXTERNO
        CONSULTA = "INSERT INTO empleados_fexternos(ID_FEXTERNO,NOMBRE,DIRECCION," & _
                   "COD_POSTAL,CIF,TELEFONO," & _
                   "FAX,RESPONSABLE,CARGO," & _
                   "TIPO,EMAIL," & _
                   "OBSERVACIONES,WEB," & _
                   "MUNICIPIO_ID,PROVINCIA_ID,PAIS_ID,ADJUNTOS)" & _
                   "values(" & _
                   ID_FEXTERNO & ",'" & NOMBRE & "','" & DIRECCION & "'," & _
                   COD_POSTAL & ",'" & CIF & "','" & TELEFONO & "','" & _
                   FAX & "','" & responsable & "','" & CARGO & "'," & _
                   IIf(Trim(tipo) = "", "0", "1") & ",'" & EMAIL & "','" & _
                   OBSERVACIONES & "','" & WEB & "'," & _
                   MUNICIPIO_ID & "," & PROVINCIA_ID & "," & PAIS_ID & ",'" & ADJUNTOS & "')"
        execute_bd CONSULTA
    End If
    Insertar = True
    Exit Function
fallo:
    Insertar = False
    MsgBox "Error al guardar el formador en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Sub CrearID_FEXTERNO()
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(ID_FEXTERNO) FROM empleados_fexternos"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_FEXTERNO = 1
    Else
        ID_FEXTERNO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Modificar() As Boolean
    Dim CONSULTA As String
    On Error GoTo fallo:
    CONSULTA = "UPDATE empleados_fexternos SET " & _
               "NOMBRE = '" & NOMBRE & "'" & _
               ",DIRECCION = '" & DIRECCION & "',COD_POSTAL = " & COD_POSTAL & _
               ",CIF = '" & CIF & "',TELEFONO = '" & TELEFONO & "'" & _
               ",FAX = '" & FAX & "',RESPONSABLE = '" & responsable & "'" & _
               ",CARGO = '" & CARGO & "',TIPO = '" & tipo & "'" & _
               ",EMAIL = '" & EMAIL & "'" & _
               ",OBSERVACIONES = '" & OBSERVACIONES & "'" & _
               ",WEB = '" & WEB & "'" & _
               ",ADJUNTOS = '" & ADJUNTOS & "'" & _
               ",MUNICIPIO_ID = " & MUNICIPIO_ID & ",PROVINCIA_ID = " & PROVINCIA_ID & _
               ",PAIS_ID = " & PAIS_ID & _
               " WHERE " & _
               " ID_FEXTERNO = " & ID_FEXTERNO
               execute_bd CONSULTA
    Modificar = True
    MsgBox "Formador modificado correctamente. Código : " & ID_FEXTERNO, vbInformation, App.Title
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el formador en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Eliminar() As Boolean
    Dim CONSULTA As String
    Dim rs As ADODB.Recordset
    On Error GoTo fallo:
    CONSULTA = "SELECT * FROM EMPLEADOS_CUALIFICACIONES WHERE EMPLEADO_ID_FORMADOR = " & ID_FEXTERNO & " AND MODALIDAD = 1"
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount <> 0 Then
        MsgBox "Existen empleados con cualificaciones del formador seleccionado. No se puede eliminar.", vbCritical, App.Title
        Exit Function
    End If
    Set rs = Nothing
    CONSULTA = "DELETE FROM empleados_fexternos " & _
               " WHERE " & _
               " ID_FEXTERNO = " & ID_FEXTERNO
               execute_bd CONSULTA
    Eliminar = True
    MsgBox "Formador eliminado correctamente. Código : " & ID_FEXTERNO, vbInformation, App.Title
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el Formador de la bd : " & Err.Description, vbCritical, Err.Number
End Function

Public Function Listado(lNombre As String, lCif As String, lTelefono As String) As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * " & _
               "  FROM empleados_fexternos " & _
               " WHERE ID_FEXTERNO <> 0 " & _
               "   AND NOMBRE LIKE '%" & lNombre & "%' " & _
               "   AND CIF LIKE '%" & lCif & "%' " & _
               "   AND TELEFONO LIKE '%" & lTelefono & "%' " & _
               " ORDER BY NOMBRE"
    Set Listado = datos_bd(CONSULTA)
End Function

Public Function duplicado() As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_FEXTERNO FROM empleados_fexternos WHERE NOMBRE = '" & NOMBRE & "'"
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount = 0 Then
        duplicado = False
    Else
        ID_FEXTERNO = rs(0)
        duplicado = True
    End If
    Exit Function
fallo:
    duplicado = False
End Function

Public Function Listado_Combo() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_FEXTERNO,NOMBRE FROM empleados_fexternos order by nombre"
    Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "empleados_fexternos"
        .setDESCRIPCION = "empleados_fexternos"
        .setPK = "ID_FEXTERNO"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
