VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeterminaciones_conf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDETERMINACIONES_CONF
'***************************************************************
Private DETERMINACION_ID As Long
Private rango As String
Private DIF_DUPLICADOS As String
Private DIF_DUPLICADOS_NUMERICA As String
Private DIF_AVISO As String
Private INCERTIDUMBRE As String
Private DIF_HISTORICO As String
Private C_VARIACION As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDETERMINACIONES_CONF
'***************************************************************
Public Property Let setDETERMINACION_ID(ByVal dato As Long)
        DETERMINACION_ID = dato
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
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDETERMINACIONES_CONF
'***************************************************************
Public Property Get getDETERMINACION_ID() As Long
        getDETERMINACION_ID = DETERMINACION_ID
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
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDETERMINACIONES_CONF
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DETERMINACIONES_CONF WHERE DETERMINACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DETERMINACION_ID = rs("DETERMINACION_ID")
                rango = rs("RANGO")
                DIF_DUPLICADOS = rs("DIF_DUPLICADOS")
                DIF_DUPLICADOS_NUMERICA = rs("DIF_DUPLICADOS_NUMERICA")
                DIF_AVISO = rs("DIF_AVISO")
                INCERTIDUMBRE = rs("INCERTIDUMBRE")
                DIF_HISTORICO = rs("DIF_HISTORICO")
                C_VARIACION = rs("C_VARIACION")
                
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDeterminaciones_conf)", vbCritical, Err.Description
End Function
Public Function Insertar(DETERMINACION_ID As Long, Optional ORDEN As Integer = -1) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE into determinaciones_conf " & _
                   " select a.ID_DETERMINACION,g.RANGO,g.DIF_DUPLICADOS,g.DIF_DUPLICADOS_NUMERICA,g.DIF_AVISO,g.INCERTIDUMBRE,g.DIF_HISTORICO,g.C_VARIACION " & _
                   " from determinaciones a, tipos_determinacion_conf g " & _
                   " where a.TIPO_DETERMINACION_ID = g.TIPO_DETERMINACION_ID " & _
                   "   and a.id_determinacion = " & DETERMINACION_ID & _
                   IIf(ORDEN >= 0, " and g.ORDEN = " & ORDEN, "")
        execute_bd consulta
        Insertar = True
        Exit Function
fallo:
        Insertar = False
        MsgBox "Error al insertar (clsDeterminaciones_conf)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DETERMINACIONES_CONF " & _
                "    WHERE DETERMINACION_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDeterminaciones_conf)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DETERMINACIONES_CONF ORDER BY DETERMINACION_ID"
        Set Listado = datos_bd(consulta)
End Function
