VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSNC
'***************************************************************
Private ID_NC As Long
Private NUMERO As Long
Private TIPO_HECHO_ID As Long
Private DEPARTAMENTO_ID As Long
Private fecha As String
Private DESCRIPCION As String
Private ORIGEN_ID As Long
Private AFECTADO_ID As Long
Private ACCION_INMEDIATA As String
Private ANALISIS_CAUSAS As String
Private EVALUADA As Long
Private NO_PROCEDENTE As Long
Private IMPACTO As Long
Private Evaluacion As String
Private IMPACTO_TEXTO As String
Private fecha_cierre As String
Private USUARIO_ID As Long
Private ESTADO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSNC
'***************************************************************
Public Property Let setID_NC(ByVal dato As Long)
        ID_NC = dato
End Property
Public Property Let setNUMERO(ByVal dato As Long)
        NUMERO = dato
End Property
Public Property Let setTIPO_HECHO_ID(ByVal dato As Long)
        TIPO_HECHO_ID = dato
End Property
Public Property Let setDEPARTAMENTO_ID(ByVal dato As Long)
        DEPARTAMENTO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setORIGEN_ID(ByVal dato As Long)
        ORIGEN_ID = dato
End Property
Public Property Let setAFECTADO_ID(ByVal dato As Long)
        AFECTADO_ID = dato
End Property
Public Property Let setACCION_INMEDIATA(ByVal dato As String)
        ACCION_INMEDIATA = dato
End Property
Public Property Let setANALISIS_CAUSAS(ByVal dato As String)
        ANALISIS_CAUSAS = dato
End Property
Public Property Let setEVALUADA(ByVal dato As Long)
        EVALUADA = dato
End Property
Public Property Let setNO_PROCEDENTE(ByVal dato As Long)
        NO_PROCEDENTE = dato
End Property
Public Property Let setIMPACTO(ByVal dato As Long)
        IMPACTO = dato
End Property
Public Property Let setEVALUACION(ByVal dato As String)
        Evaluacion = dato
End Property
Public Property Let setIMPACTO_TEXTO(ByVal dato As String)
        IMPACTO_TEXTO = dato
End Property
Public Property Let setFECHA_CIERRE(ByVal dato As String)
        fecha_cierre = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Long)
        ESTADO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSNC
'***************************************************************
Public Property Get getID_NC() As Long
        getID_NC = ID_NC
End Property
Public Property Get getNUMERO() As Long
        getNUMERO = NUMERO
End Property
Public Property Get getTIPO_HECHO_ID() As Long
        getTIPO_HECHO_ID = TIPO_HECHO_ID
End Property
Public Property Get getDEPARTAMENTO_ID() As Long
        getDEPARTAMENTO_ID = DEPARTAMENTO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getORIGEN_ID() As Long
        getORIGEN_ID = ORIGEN_ID
End Property
Public Property Get getAFECTADO_ID() As Long
        getAFECTADO_ID = AFECTADO_ID
End Property
Public Property Get getACCION_INMEDIATA() As String
        getACCION_INMEDIATA = ACCION_INMEDIATA
End Property
Public Property Get getANALISIS_CAUSAS() As String
        getANALISIS_CAUSAS = ANALISIS_CAUSAS
End Property
Public Property Get getEVALUADA() As Long
        getEVALUADA = EVALUADA
End Property
Public Property Get getNO_PROCEDENTE() As Long
        getNO_PROCEDENTE = NO_PROCEDENTE
End Property
Public Property Get getIMPACTO() As Long
        getIMPACTO = IMPACTO
End Property
Public Property Get getEVALUACION() As String
        getEVALUACION = Evaluacion
End Property
Public Property Get getIMPACTO_TEXTO() As String
        getIMPACTO_TEXTO = IMPACTO_TEXTO
End Property
Public Property Get getFECHA_CIERRE() As String
        getFECHA_CIERRE = fecha_cierre
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getESTADO_ID() As Long
        getESTADO_ID = ESTADO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSNC
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM NC WHERE ID_NC = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_NC = rs("ID_NC")
                NUMERO = rs("NUMERO")
                TIPO_HECHO_ID = rs("TIPO_HECHO_ID")
                DEPARTAMENTO_ID = rs("DEPARTAMENTO_ID")
                fecha = rs("FECHA")
                DESCRIPCION = rs("DESCRIPCION")
                ORIGEN_ID = rs("ORIGEN_ID")
                AFECTADO_ID = rs("AFECTADO_ID")
                ACCION_INMEDIATA = rs("ACCION_INMEDIATA")
                ANALISIS_CAUSAS = rs("ANALISIS_CAUSAS")
                EVALUADA = rs("EVALUADA")
                NO_PROCEDENTE = rs("NO_PROCEDENTE")
                IMPACTO = rs("IMPACTO")
                Evaluacion = rs("EVALUACION")
                IMPACTO_TEXTO = rs("IMPACTO_TEXTO")
                fecha_cierre = rs("FECHA_CIERRE")
                USUARIO_ID = rs("USUARIO_ID")
                ESTADO_ID = rs("ESTADO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsNc)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_NC) FROM NC"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_NC = 1
        Else
                ID_NC = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Calcular_Numero(tipo As Long) As Long
        Dim rs As ADODB.Recordset
        Dim consulta As String
        On Error GoTo fallo
        consulta = "SELECT MAX(NUMERO) FROM NC " & _
                   " WHERE YEAR(FECHA) = " & Year(Date) & _
                   "   AND TIPO_HECHO_ID = " & tipo
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            Calcular_Numero = 1
        Else
            Calcular_Numero = rs.Fields(0) + 1
        End If
        Set rs = Nothing
        Exit Function
fallo:
        MsgBox "Error en Calcular_Numero de la NC : " & Err.Description, vbCritical, App.Title
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_NC = 0 Then
            Me.CrearID
        End If
        consulta = "INSERT INTO NC " & _
                   "  VALUES (" & _
                        ID_NC & "," & NUMERO & "," & TIPO_HECHO_ID & "," & _
                        DEPARTAMENTO_ID & "," & "'" & fecha & "'" & "," & "'" & DESCRIPCION & "'" & "," & _
                        ORIGEN_ID & "," & AFECTADO_ID & "," & "'" & ACCION_INMEDIATA & "'" & "," & "'" & ANALISIS_CAUSAS & "'" & "," & _
                        EVALUADA & "," & NO_PROCEDENTE & "," & IMPACTO & "," & _
                        "'" & Evaluacion & "'" & "," & "'" & IMPACTO_TEXTO & "'" & "," & "'" & fecha_cierre & "'" & "," & _
                        USUARIO_ID & "," & ESTADO_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_NC
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsNc)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE NC SET " & _
                        " NUMERO = " & NUMERO & "," & _
                        " TIPO_HECHO_ID = " & TIPO_HECHO_ID & "," & _
                        " DEPARTAMENTO_ID = " & DEPARTAMENTO_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " ORIGEN_ID = " & ORIGEN_ID & "," & _
                        " AFECTADO_ID = " & AFECTADO_ID & "," & _
                        " ACCION_INMEDIATA = '" & ACCION_INMEDIATA & "'," & _
                        " ANALISIS_CAUSAS = '" & ANALISIS_CAUSAS & "'," & _
                        " EVALUADA = " & EVALUADA & "," & _
                        " NO_PROCEDENTE = " & NO_PROCEDENTE & "," & _
                        " IMPACTO = " & IMPACTO & "," & _
                        " EVALUACION = '" & Evaluacion & "'," & _
                        " IMPACTO_TEXTO = '" & IMPACTO_TEXTO & "'," & _
                        " FECHA_CIERRE = '" & fecha_cierre & "'," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " ESTADO_ID = " & ESTADO_ID & "" & _
                " WHERE ID_NC = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsNc)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM NC " & _
                "    WHERE ID_NC = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsNc)", vbCritical, Err.Description
End Function
Public Function Listado(tipo As String, origen As String, ESTADO As String, nombre As String, afectado As String) As ADODB.Recordset
        Dim consulta As String
        Dim oTIPO As String
        Dim oorigen As String
        Dim oestado As String
        Dim oafectado As String
        If tipo <> "0" Then
            oTIPO = " AND A.TIPO_HECHO_ID = " & CInt(tipo)
        End If
        If origen <> "0" Then
            oorigen = " AND A.ORIGEN_ID = " & CInt(origen)
        End If
        If ESTADO <> "0" Then
            oestado = " AND A.ESTADO_ID = " & CInt(ESTADO)
        End If
        If afectado <> "0" Then
            oafectado = " AND A.AFECTADO_ID = " & CInt(afectado)
        End If
        consulta = "SELECT A.ID_NC,A.NUMERO,B.DESCRIPCION,C.DESCRIPCION,A.DESCRIPCION,A.FECHA,A.FECHA_CIERRE,D.DESCRIPCION,E.DESCRIPCION " & _
                   "  FROM NC A, decodificadora B, decodificadora C, decodificadora D, decodificadora E " & _
                   " WHERE (A.TIPO_HECHO_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.NC_TIPOS_HECHOS & ")" & _
                   "   AND (A.ORIGEN_ID = C.VALOR AND C.CODIGO = " & DECODIFICADORA.NC_ORIGENES & ")" & _
                   "   AND (A.ESTADO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.NC_ESTADOS & ")" & _
                   "   AND (A.AFECTADO_ID = E.VALOR AND E.CODIGO = " & DECODIFICADORA.NC_AFECTADO & ")" & _
                   "   AND A.DESCRIPCION LIKE '%" & nombre & "%' " & _
                   oTIPO & _
                   oorigen & _
                   oestado & _
                   oafectado & _
                   " ORDER BY A.ID_NC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_Codigo(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_NC,A.NUMERO,B.DESCRIPCION,C.DESCRIPCION,A.DESCRIPCION,A.FECHA,A.FECHA_CIERRE,D.DESCRIPCION,E.DESCRIPCION " & _
                   "  FROM NC A, decodificadora B, decodificadora C, decodificadora D, decodificadora E " & _
                   " WHERE (A.TIPO_HECHO_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.NC_TIPOS_HECHOS & ")" & _
                   "   AND (A.ORIGEN_ID = C.VALOR AND C.CODIGO = " & DECODIFICADORA.NC_ORIGENES & ")" & _
                   "   AND (A.ESTADO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.NC_ESTADOS & ")" & _
                   "   AND (A.AFECTADO_ID = E.VALOR AND E.CODIGO = " & DECODIFICADORA.NC_AFECTADO & ")" & _
                   "   AND A.ID_NC = " & ID
        Set Listado_por_Codigo = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_NC,NUMERO FROM NC ORDER BY NUMERO"
        Set Listado_Combo = datos_bd(consulta)
End Function

