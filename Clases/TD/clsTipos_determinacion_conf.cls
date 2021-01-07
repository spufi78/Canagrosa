VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_determinacion_conf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_DETERMINACION_CONF
'***************************************************************
Private TIPO_DETERMINACION_ID As Long
Private ORDEN As Long
Private rango As String
Private DIF_DUPLICADOS As String
Private DIF_DUPLICADOS_NUMERICA As String
Private DIF_AVISO As String
Private INCERTIDUMBRE As String
Private DIF_HISTORICO As String
Private C_VARIACION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_DETERMINACION_CONF
'***************************************************************
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setRANGO(ByVal dato As String)
        rango = dato
End Property
Public Property Let setDIF_DUPLICADOS(ByVal dato As String)
        DIF_DUPLICADOS = dato
End Property
Public Property Let setDIF_DUPLICADOS_NUMERICA(ByVal dato As String)
        DIF_DUPLICADOS_NUMERICA = dato
End Property
Public Property Let setDIF_AVISO(ByVal dato As String)
        DIF_AVISO = dato
End Property
Public Property Let setINCERTIDUMBRE(ByVal dato As String)
        INCERTIDUMBRE = dato
End Property
Public Property Let setDIF_HISTORICO(ByVal dato As String)
        DIF_HISTORICO = dato
End Property
Public Property Let setC_VARIACION(ByVal dato As String)
        C_VARIACION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_DETERMINACION_CONF
'***************************************************************
Public Property Get getTIPO_DETERMINACION_ID() As Long
        getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getRANGO() As String
        getRANGO = rango
End Property
Public Property Get getDIF_DUPLICADOS() As String
        getDIF_DUPLICADOS = DIF_DUPLICADOS
End Property
Public Property Get getDIF_DUPLICADOS_NUMERICA() As String
        getDIF_DUPLICADOS_NUMERICA = DIF_DUPLICADOS_NUMERICA
End Property
Public Property Get getDIF_AVISO() As String
        getDIF_AVISO = DIF_AVISO
End Property
Public Property Get getINCERTIDUMBRE() As String
        getINCERTIDUMBRE = INCERTIDUMBRE
End Property
Public Property Get getDIF_HISTORICO() As String
        getDIF_HISTORICO = DIF_HISTORICO
End Property
Public Property Get getC_VARIACION() As String
        getC_VARIACION = C_VARIACION
End Property
Public Function CARGAR(TIPO_DETERMINACION_ID As Long, Optional ORDEN As Integer = -1) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_DETERMINACION_CONF " & _
               " WHERE TIPO_DETERMINACION_ID = " & TIPO_DETERMINACION_ID & _
               IIf(ORDEN >= 0, " AND ORDEN = " & ORDEN, "") & _
               " ORDER BY ORDEN LIMIT 1"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        CARGAR = False
    Else
        CARGAR = True
        TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
        ORDEN = rs("ORDEN")
        rango = rs("RANGO")
        DIF_DUPLICADOS = texto(rs("DIF_DUPLICADOS"))
        DIF_DUPLICADOS_NUMERICA = texto(rs("DIF_DUPLICADOS_NUMERICA"))
        DIF_AVISO = texto(rs("DIF_AVISO"))
        INCERTIDUMBRE = texto(rs("INCERTIDUMBRE"))
        DIF_HISTORICO = texto(rs("DIF_HISTORICO"))
        C_VARIACION = texto(rs("C_VARIACION"))
    End If
    Set rs = Nothing
    Exit Function
fallo:
    CARGAR = False
    MsgBox "Error al cargar la configuración por defecto (CargarDefecto)", vbCritical, Err.Description
End Function


Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TIPOS_DETERMINACION_CONF " & _
                   "  VALUES (" & _
                        TIPO_DETERMINACION_ID & "," & ORDEN & "," & "'" & rango & "'" & "," & _
                        numericoBD(DIF_DUPLICADOS) & "," & "'" & DIF_DUPLICADOS_NUMERICA & "'" & "," & numericoBD(DIF_AVISO) & "," & _
                        numericoBD(INCERTIDUMBRE) & "," & _
                        numericoBD(DIF_HISTORICO) & "," & _
                        numericoBD(C_VARIACION) & _
                ")"
        execute_bd consulta
        Insertar = TIPO_DETERMINACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTipos_determinacion_conf)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_DETERMINACION_CONF " & _
                "    WHERE TIPO_DETERMINACION_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTipos_determinacion_conf)", vbCritical, Err.Description
End Function
Public Function Listado(TIPO_DETERMINACION_ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_DETERMINACION_CONF WHERE TIPO_DETERMINACION_ID = " & TIPO_DETERMINACION_ID & " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function


