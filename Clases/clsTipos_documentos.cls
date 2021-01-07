VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_TIPO_DOCUMENTO As Integer
Private nombre As String
Private PLANTILLA As String
Private fecha As String
Private CONVERTIDO As Integer
Private MANUAL As Integer
Private parametros As String
Private FIRMA_CANAGROSA As Integer
Private FIRMA_RESPONSABLE As Integer
Private FIRMA1_X1 As Integer
Private FIRMA1_Y1 As Integer
Private FIRMA1_X2 As Integer
Private FIRMA1_Y2 As Integer
Private FIRMA2_X1 As Integer
Private FIRMA2_Y1 As Integer
Private FIRMA2_X2 As Integer
Private FIRMA2_Y2 As Integer
Public Property Let setID_TIPO_DOCUMENTO(ByVal dato As Integer)
    ID_TIPO_DOCUMENTO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setPLANTILLA(ByVal dato As String)
    PLANTILLA = Trim(dato)
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = Trim(dato)
End Property
Public Property Let setCONVERTIDO(ByVal dato As Integer)
    CONVERTIDO = dato
End Property
Public Property Let setMANUAL(ByVal dato As Integer)
    MANUAL = dato
End Property
Public Property Let setPARAMETROS(ByVal dato As String)
    parametros = dato
End Property
Public Property Let setFIRMA_CANAGROSA(ByVal dato As Integer)
    FIRMA_CANAGROSA = dato
End Property
Public Property Let setFIRMA_RESPONSABLE(ByVal dato As Integer)
    FIRMA_RESPONSABLE = dato
End Property
Public Property Let setFIRMA1_X1(ByVal dato As Integer)
    FIRMA1_X1 = dato
End Property
Public Property Let setFIRMA1_Y1(ByVal dato As Integer)
    FIRMA1_Y1 = dato
End Property
Public Property Let setFIRMA1_X2(ByVal dato As Integer)
    FIRMA1_X2 = dato
End Property
Public Property Let setFIRMA1_Y2(ByVal dato As Integer)
    FIRMA1_Y2 = dato
End Property
Public Property Let setFIRMA2_X1(ByVal dato As Integer)
    FIRMA2_X1 = dato
End Property
Public Property Let setFIRMA2_Y1(ByVal dato As Integer)
    FIRMA2_Y1 = dato
End Property
Public Property Let setFIRMA2_X2(ByVal dato As Integer)
    FIRMA2_X2 = dato
End Property
Public Property Let setFIRMA2_Y2(ByVal dato As Integer)
    FIRMA2_Y2 = dato
End Property


Public Property Get getID_TIPO_DOCUMENTO() As Integer
    getID_TIPO_DOCUMENTO = ID_TIPO_DOCUMENTO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getPLANTILLA() As String
    getPLANTILLA = PLANTILLA
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Property Get getCONVERTIDO() As Integer
    getCONVERTIDO = CONVERTIDO
End Property
Public Property Get getMANUAL() As Integer
    getMANUAL = MANUAL
End Property
Public Property Get getPARAMETROS() As String
    getPARAMETROS = parametros
End Property
Public Property Get getFIRMA_CANAGROSA() As Integer
    getFIRMA_CANAGROSA = FIRMA_CANAGROSA
End Property
Public Property Get getFIRMA_RESPONSABLE() As Integer
    getFIRMA_RESPONSABLE = FIRMA_RESPONSABLE
End Property
Public Property Get getFIRMA1_X1() As Integer
    getFIRMA1_X1 = FIRMA1_X1
End Property
Public Property Get getFIRMA1_Y1() As Integer
    getFIRMA1_Y1 = FIRMA1_Y1
End Property
Public Property Get getFIRMA1_X2() As Integer
    getFIRMA1_X2 = FIRMA1_X2
End Property
Public Property Get getFIRMA1_Y2() As Integer
    getFIRMA1_Y2 = FIRMA1_Y2
End Property
Public Property Get getFIRMA2_X1() As Integer
    getFIRMA2_X1 = FIRMA2_X1
End Property
Public Property Get getFIRMA2_Y1() As Integer
    getFIRMA2_Y1 = FIRMA2_Y1
End Property
Public Property Get getFIRMA2_X2() As Integer
    getFIRMA2_X2 = FIRMA2_X2
End Property
Public Property Get getFIRMA2_Y2() As Integer
    getFIRMA2_Y2 = FIRMA2_Y2
End Property

Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_DOCUMENTOS order by NOMBRE"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_DOCUMENTO,NOMBRE FROM TIPOS_DOCUMENTOS order by NOMBRE"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function CARGAR(TIPO_DOCUMENTO As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_DOCUMENTOS WHERE ID_TIPO_DOCUMENTO=" & TIPO_DOCUMENTO
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID_TIPO_DOCUMENTO = rs("ID_TIPO_DOCUMENTO")
        nombre = rs("NOMBRE")
        PLANTILLA = rs("PLANTILLA")
        fecha = rs("FECHA")
        CONVERTIDO = rs("CONVERTIDO")
        MANUAL = rs("MANUAL")
        parametros = rs("PARAMETROS")
        FIRMA_CANAGROSA = rs("FIRMA_CANAGROSA")
        FIRMA_RESPONSABLE = rs("FIRMA_RESPONSABLE")
        FIRMA1_X1 = rs("FIRMA1_X1")
        FIRMA1_Y1 = rs("FIRMA1_Y1")
        FIRMA1_X2 = rs("FIRMA1_X2")
        FIRMA1_Y2 = rs("FIRMA1_Y2")
        FIRMA2_X1 = rs("FIRMA2_X1")
        FIRMA2_Y1 = rs("FIRMA2_Y1")
        FIRMA2_X2 = rs("FIRMA2_X2")
        FIRMA2_Y2 = rs("FIRMA2_Y2")
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
End Function
Public Function Nuevo_Formato(TIPO_DOCUMENTO As Long, Optional ID_MUESTRA As Long = 0) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.Recordset, rs2 As ADODB.Recordset
    
    consulta = "SELECT CONVERTIDO,FECHA FROM TIPOS_DOCUMENTOS WHERE ID_TIPO_DOCUMENTO=" & TIPO_DOCUMENTO
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
       Nuevo_Formato = True
    Else
        If rs(0) = 1 Then ' CONVERTIDO A CRYSTAL REPORTS
'            If ID_MUESTRA > 0 Then
'                consulta = "SELECT FECHA_CIERRE FROM muestras where id_muestra = " & ID_MUESTRA
'                Set rs2 = datos_bd(consulta)
'                If rs2.RecordCount <> 0 Then
'                    rs2.MoveFirst
'                    If IsNull(rs2("FECHA_CIERRE")) Then
'                        Nuevo_Formato = True
'                    ElseIf Format(CDate(rs2("FECHA_CIERRE")), "YYYY-MM-DD") >= Format(rs(1), "YYYY-MM-DD") Then
'                        Nuevo_Formato = True
'                    Else
'                        ' Se cerro a partir del formato antiguo
'                        Nuevo_Formato = False
'                    End If
'                End If
'            Else
                Nuevo_Formato = True
'            End If
        Else
            Nuevo_Formato = False
        End If
    End If
    Set rs = Nothing
    Exit Function
fallo:
    Nuevo_Formato = True
    MsgBox "Error al obtener el tipo de documento de la muestra.", vbCritical, Err.Description
End Function
'M1051-I
'Public Function esManual(codmuestra As Long) As Boolean
'    esManual = False
'    On Error GoTo fallo
'    Dim consulta As String
'    Dim rs As ADODB.RecordSet
'    Dim oMuestra As New clsMuestra
'    consulta = "SELECT MANUAL FROM TIPOS_DOCUMENTOS WHERE ID_TIPO_DOCUMENTO=" & oMuestra.obtener_tipo_documento(codmuestra)
'    Set rs = datos_bd(consulta)
'    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
'        esManual = False
'    Else
'        esManual = False
'        If rs(0) = 1 Then
'            If Not oMuestra.esCerrada(codmuestra) Then
'                esManual = True
'            End If
'        End If
'    End If
'    Set oMuestra = Nothing
'    Set rs = Nothing
'    Exit Function
'fallo:
'    esManual = False
'    MsgBox "Error al obtener el tipo de documento de la muestra.", vbCritical, Err.Description
'End Function
'M1051-F
Public Function ParametrosDocumento(codmuestra As Long) As String
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim oMuestra As New clsMuestra
    consulta = "SELECT PARAMETROS FROM TIPOS_DOCUMENTOS WHERE ID_TIPO_DOCUMENTO=" & oMuestra.obtener_tipo_documento(codmuestra)
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ParametrosDocumento = ""
    Else
        ParametrosDocumento = rs(0)
    End If
    log ("ParametrosDocumento : " & ParametrosDocumento)
    Set oMuestra = Nothing
    Set rs = Nothing
    Exit Function
fallo:
    ParametrosDocumento = ""
    log "Error al obtener los parametros del documento de la muestra : " & Err.Description
End Function
