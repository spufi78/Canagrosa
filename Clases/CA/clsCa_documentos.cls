VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_DOCUMENTOS
'***************************************************************
Private ID_DOCUMENTO As Long
Private FAMILIA_ID As Long
Private SUBFAMILIA_ID As Long
Private nombre As String
Private CODIGO As String
Private ENAC As Integer
Private nadcap As Integer
Private mtl As Integer
Private EQA As Integer
Private EDICION As String
Private fecha As String
Private OBSERVACIONES As String
Private RESPONSABLE_ID As Integer
Private ESTADO_ID As Integer
Private RUTA As String
Private uso As Integer
Private PLANTILLA_ID As Integer
Private NO_EDICION As Integer
Private DOCUMENTO_VINCULADO As Integer
Private COPIA_CONTROLADA As Integer
Private COPIA_LABORATORIO As String
Private COPIA_NUMERO As String
Private FORMACION As Integer
Private PLAN_ID As Integer
Private FAMILIA_REQ_ID As Integer
Private anulado As Integer
'BUG-XXXX-I
Private VINCULADO As Integer
'BUG-XXXX-F
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_DOCUMENTOS
'***************************************************************
Public Property Let setID_DOCUMENTO(ByVal dato As Long)
        ID_DOCUMENTO = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Long)
        FAMILIA_ID = dato
End Property
Public Property Let setSUBFAMILIA_ID(ByVal dato As Long)
        SUBFAMILIA_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setENAC(ByVal dato As Integer)
        ENAC = dato
End Property
Public Property Let setNADCAP(ByVal dato As Integer)
        nadcap = dato
End Property
Public Property Let setMTL(ByVal dato As Integer)
        mtl = dato
End Property
Public Property Let setEQA(ByVal dato As Integer)
        EQA = dato
End Property
Public Property Let setEDICION(ByVal dato As String)
        EDICION = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Integer)
        RESPONSABLE_ID = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Integer)
        ESTADO_ID = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        RUTA = Replace(dato, "\", "/")
End Property
Public Property Let setUSO(ByVal dato As Integer)
        uso = dato
End Property
Public Property Let setPLANTILLA_ID(ByVal dato As Integer)
        PLANTILLA_ID = dato
End Property
Public Property Let setNO_EDICION(ByVal dato As Integer)
        NO_EDICION = dato
End Property
Public Property Let setDOCUMENTO_VINCULADO(ByVal dato As Integer)
        DOCUMENTO_VINCULADO = dato
End Property
Public Property Let setCOPIA_CONTROLADA(ByVal dato As Integer)
    COPIA_CONTROLADA = dato
End Property
Public Property Let setCOPIA_LABORATORIO(ByVal dato As String)
    COPIA_LABORATORIO = dato
End Property
Public Property Let setCOPIA_NUMERO(ByVal dato As String)
    COPIA_NUMERO = dato
End Property
Public Property Let setFORMACION(ByVal dato As Integer)
    FORMACION = dato
End Property
Public Property Let setPLAN_ID(ByVal dato As Integer)
    PLAN_ID = dato
End Property
Public Property Let setFAMILIA_REQ_ID(ByVal dato As Integer)
    FAMILIA_REQ_ID = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
'BUG-XXXX-I
Public Property Let setVINCULADO(ByVal dato As Integer)
    VINCULADO = dato
End Property
'BUG-XXXX-F

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_DOCUMENTOS
'***************************************************************
Public Property Get getID_DOCUMENTO() As Long
        getID_DOCUMENTO = ID_DOCUMENTO
End Property
Public Property Get getFAMILIA_ID() As Long
        getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getSUBFAMILIA_ID() As Long
        getSUBFAMILIA_ID = SUBFAMILIA_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getENAC() As Integer
        getENAC = ENAC
End Property
Public Property Get getNADCAP() As Integer
        getNADCAP = nadcap
End Property
Public Property Get getMTL() As Integer
        getMTL = mtl
End Property
Public Property Get getEQA() As Integer
        getEQA = EQA
End Property
Public Property Get getEDICION() As String
        getEDICION = EDICION
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getRESPONSABLE_ID() As Integer
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getESTADO_ID() As Integer
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getRUTA() As String
        getRUTA = Replace(RUTA, "/", "\")
End Property
Public Property Get getUSO() As Integer
        getUSO = uso
End Property
Public Property Get getPLANTILLA_ID() As Integer
        getPLANTILLA_ID = PLANTILLA_ID
End Property
Public Property Get getNO_EDICION() As Integer
        getNO_EDICION = NO_EDICION
End Property
Public Property Get getDOCUMENTO_VINCULADO() As Integer
        getDOCUMENTO_VINCULADO = DOCUMENTO_VINCULADO
End Property
Public Property Get getCOPIA_CONTROLADA() As Integer
    getCOPIA_CONTROLADA = COPIA_CONTROLADA
End Property
Public Property Get getCOPIA_LABORATORIO() As String
    getCOPIA_LABORATORIO = COPIA_LABORATORIO
End Property
Public Property Get getCOPIA_NUMERO() As String
    getCOPIA_NUMERO = COPIA_NUMERO
End Property
Public Property Get getFORMACION() As Integer
    getFORMACION = FORMACION
End Property
Public Property Get getPLAN_ID() As Integer
    getPLAN_ID = PLAN_ID
End Property
Public Property Get getFAMILIA_REQ_ID() As Integer
    getFAMILIA_REQ_ID = FAMILIA_REQ_ID
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
'BUG-XXXX-I
Public Property Get getVINCULADO() As Integer
    getVINCULADO = VINCULADO
End Property
'BUG-XXXX-F
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_DOCUMENTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ca_documentos WHERE ID_DOCUMENTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_DOCUMENTO = rs("ID_DOCUMENTO")
                FAMILIA_ID = rs("FAMILIA_ID")
                SUBFAMILIA_ID = rs("SUBFAMILIA_ID")
                
                nombre = rs("NOMBRE")
                CODIGO = rs("CODIGO")
                ENAC = rs("ENAC")
                nadcap = rs("NADCAP")
                mtl = rs("MTL")
                EQA = rs("EQA")
                EDICION = rs("EDICION")
                fecha = rs("FECHA")
                OBSERVACIONES = rs("OBSERVACIONES")
                RESPONSABLE_ID = rs("RESPONSABLE_ID")
                ESTADO_ID = rs("ESTADO_ID")
                RUTA = rs("RUTA")
                uso = rs("USO")
                PLANTILLA_ID = rs("PLANTILLA_ID")
                NO_EDICION = rs("NO_EDICION")
                DOCUMENTO_VINCULADO = rs("DOCUMENTO_VINCULADO")
                COPIA_CONTROLADA = rs("COPIA_CONTROLADA")
                COPIA_LABORATORIO = rs("COPIA_LABORATORIO")
                COPIA_NUMERO = rs("COPIA_NUMERO")
                FORMACION = rs("FORMACION")
                PLAN_ID = rs("PLAN_ID")
                FAMILIA_REQ_ID = rs("FAMILIA_REQ_ID")
                anulado = rs("ANULADO")
'BUG-XXXX-I
                VINCULADO = rs("VINCULADO")
'BUG-XXXX-F
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_documentos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_DOCUMENTO) FROM CA_DOCUMENTOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_DOCUMENTO = 1
        Else
                ID_DOCUMENTO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO ca_documentos " & _
                   "  VALUES (" & _
                        ID_DOCUMENTO & "," & FAMILIA_ID & "," & SUBFAMILIA_ID & "," & "'" & nombre & "'" & "," & _
                        "'" & CODIGO & "'" & "," & ENAC & "," & nadcap & "," & mtl & "," & EQA & "," & _
                        "'" & EDICION & "'" & "," & "'" & fecha & "'" & "," & "'" & OBSERVACIONES & "'" & "," & _
                        RESPONSABLE_ID & "," & ESTADO_ID & ",'" & RUTA & "'," & uso & "," & PLANTILLA_ID & "," & NO_EDICION & "," & DOCUMENTO_VINCULADO & "," & _
                        COPIA_CONTROLADA & "," & "'" & COPIA_LABORATORIO & "'" & ",'" & COPIA_NUMERO & "'," & FORMACION & "," & PLAN_ID & "," & FAMILIA_REQ_ID & "," & anulado & "," & VINCULADO & _
                   ")"
        execute_bd consulta
        Insertar = ID_DOCUMENTO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_documentos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ca_documentos SET " & _
                        " FAMILIA_ID = " & FAMILIA_ID & "," & _
                        " SUBFAMILIA_ID = " & SUBFAMILIA_ID & "," & _
                        " NOMBRE = '" & nombre & "'," & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " ENAC = " & ENAC & "," & _
                        " NADCAP = " & nadcap & "," & _
                        " MTL = " & mtl & "," & _
                        " EQA = " & EQA & "," & _
                        " EDICION = '" & EDICION & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " RESPONSABLE_ID = " & RESPONSABLE_ID & "," & _
                        " ESTADO_ID = " & ESTADO_ID & "," & _
                        " USO = " & uso & "," & _
                        " PLANTILLA_ID = " & PLANTILLA_ID & "," & _
                        " NO_EDICION = " & NO_EDICION & "," & _
                        " DOCUMENTO_VINCULADO = " & DOCUMENTO_VINCULADO & "," & _
                        " COPIA_CONTROLADA = " & COPIA_CONTROLADA & "," & _
                        " COPIA_LABORATORIO = '" & COPIA_LABORATORIO & "'," & _
                        " FORMACION = " & FORMACION & "," & _
                        " PLAN_ID = " & PLAN_ID & "," & _
                        " FAMILIA_REQ_ID = " & FAMILIA_REQ_ID & "," & _
                        " COPIA_NUMERO = '" & COPIA_NUMERO & "'" & _
                        " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_documentos)", vbCritical, Err.Description
End Function
Public Function Informar_ruta(ID As Long, sruta As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ca_documentos SET " & _
                        " RUTA = '" & Replace(sruta, "\", "/") & "'" & _
                " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        Informar_ruta = True
        Exit Function
fallo:
        Informar_ruta = False
        MsgBox "Error al informar la ruta del documento (clsCa_documentos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        execute_bd "UPDATE ca_documentos SET ANULADO = 1 WHERE ID_DOCUMENTO = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_documentos)", vbCritical, Err.Description
End Function
'BUG-XXXX-I
Public Function VINCULAR(ID As Long) As Boolean
        On Error GoTo fallo
        execute_bd "UPDATE ca_documentos SET VINCULADO = 1 WHERE ID_DOCUMENTO = " & ID & " AND VINCULADO <> 1"
        VINCULAR = True
        Exit Function
fallo:
        VINCULAR = False
        MsgBox "Error al marcar como vinculado (clsCa_documentos)", vbCritical, Err.Description
End Function

Public Function Desvincular(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
         
        Dim rs As ADODB.Recordset
 
        execute_bd "UPDATE ca_documentos SET VINCULADO = 0 WHERE ID_DOCUMENTO = " & ID & " AND VINCULADO <> 0"
        Desvincular = True
        Exit Function
fallo:
        Desvincular = False
        MsgBox "Error al desvincular (clsCa_documentos)", vbCritical, Err.Description
End Function
'BUG-XXXX-F


Public Function Listado(familia As String, subfamilia As String, ESTADO As String, nombre As String, CODIGO As String, uso As Integer, ENAC As Integer, nadcap As Integer, mtl As Integer, EQA As Integer, responsable As String, COPIAS As Integer, chkFechas As Boolean, fdesde As String, fhasta As String, chkSinTocar As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim oFamilia As String
        Dim oSUBFAMILIA As String
        Dim oestado As String
        Dim ouso As String
        Dim oENAC As String
        Dim oNADCAP As String
        Dim oMTL As String
        Dim oEQA As String
        Dim oCopias As String
        Dim oResponsable As String
        oFamilia = ""
        oestado = ""
        If familia <> "0" Then
            oFamilia = " AND A.FAMILIA_ID = " & CInt(familia)
        End If
        If subfamilia <> "0" Then
            oSUBFAMILIA = " AND A.SUBFAMILIA_ID = " & CInt(subfamilia)
        End If
        If ESTADO <> "0" Then
            oestado = " AND A.ESTADO_ID = " & CInt(ESTADO)
        End If
        If responsable <> "0" Then
            oResponsable = " AND A.RESPONSABLE_ID = " & CInt(responsable)
        End If
        If uso = 0 Then
            ouso = " AND A.USO = 1 "
        End If
        If ENAC = 1 Then
            oENAC = " AND A.ENAC = 1 "
        End If
        If nadcap = 1 Then
            oNADCAP = " AND A.NADCAP = 1 "
        End If
        If mtl = 1 Then
            oMTL = " AND A.MTL = 1 "
        End If
        If EQA = 1 Then
            oEQA = " AND A.EQA = 1 "
        End If
        If COPIAS = 1 Then
            oCopias = " AND A.COPIA_CONTROLADA = 1 "
        End If
        Dim fechas As String
'        If chkFechas = True Then
'            fechas = " AND DATE(A.FECHA) >= '" & Format(fdesde, "yyyy-mm-dd") & "' AND DATE(A.FECHA) <= '" & Format(fhasta, "yyyy-mm-dd") & "'"
'        End If
        ' M0665-I
        Dim perFamilias As String
        If USUARIO.getPER_FAMILIAS_CA = False Then
            Dim oParam As New clsParametros
            oParam.Carga parametros.PARAM_CA_FAMILIAS_NO_VISIBLES, ""
            If oParam.getVALOR <> "" Then
                perFamilias = " AND A.FAMILIA_ID NOT IN (" & oParam.getVALOR & ") "
            End If
            Set oParam = Nothing
        End If
        If chkSinTocar = True Then
            fechas = fechas & " AND (A.ESTADO_ID = " & C_CA_DOCUMENTOS_ESTADOS.CA_VIGOR & " AND DATE_SUB(CURRENT_DATE,INTERVAL 4 YEAR) > IFNULL(STR_TO_DATE(A.FECHA,'%d-%m-%Y'),A.FECHA))"
        End If
        ' M0665-F
        consulta = "SELECT A.ID_DOCUMENTO,A.NOMBRE,B.DESCRIPCION,A.CODIGO,A.EDICION,A.FECHA,C.DESCRIPCION,A.USO," & _
                   "       A.COPIA_CONTROLADA,A.COPIA_LABORATORIO,A.NO_EDICION, " & _
                   "       IF(A.ESTADO_ID = " & C_CA_DOCUMENTOS_ESTADOS.CA_VIGOR & " AND DATE_SUB(CURRENT_DATE,INTERVAL 5 YEAR) > IFNULL(STR_TO_DATE(A.FECHA,'%d-%m-%Y'),A.FECHA),1,0) AS ROJO " & _
                   "  FROM ca_documentos A, decodificadora B, decodificadora C " & _
                   " WHERE (A.FAMILIA_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS & ")" & _
                   "   AND (A.ESTADO_ID = C.VALOR AND C.CODIGO = " & DECODIFICADORA.CA_DOCUMENTOS_ESTADOS & ")" & _
                   "   AND A.NOMBRE LIKE '%" & nombre & "%' " & _
                   "   AND A.CODIGO LIKE '%" & CODIGO & "%' " & _
                   "   AND A.ANULADO = 0 " & _
                   perFamilias & _
                   oFamilia & _
                   oSUBFAMILIA & _
                   oestado & oResponsable & _
                   ouso & oENAC & oNADCAP & oMTL & oEQA & oCopias & fechas & _
                   " ORDER BY B.PARAMETROS, A.CODIGO, A.ID_DOCUMENTO"
        
        
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_Codigo(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_DOCUMENTO,A.NOMBRE,B.DESCRIPCION,A.CODIGO,A.EDICION,A.FECHA,C.DESCRIPCION,A.USO,A.COPIA_CONTROLADA,A.NO_EDICION " & _
                   "  FROM ca_documentos A, decodificadora B, decodificadora C " & _
                   " WHERE (A.FAMILIA_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS & ")" & _
                   "   AND (A.ESTADO_ID = C.VALOR AND C.CODIGO = " & DECODIFICADORA.CA_DOCUMENTOS_ESTADOS & ")" & _
                   "   AND A.ANULADO = 0 " & _
                   "   AND A.ID_DOCUMENTO = " & ID
        Set Listado_por_Codigo = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_DOCUMENTO,FAMILIA_ID FROM ca_documentos ORDER BY FAMILIA_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_Auditoria(filtro1 As String, filtro2 As String, nadcap As Boolean, mtl As Boolean, EQA As Boolean, ENAC As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim f1 As String
        Dim f2 As String
        Dim f3 As String
        If filtro1 <> "" Then
            f1 = "   AND A.FAMILIA_ID IN (" & filtro1 & ") "
        End If
        If filtro2 <> "" Then
'            If f1 = "" Then
                f2 = "   AND A.SUBFAMILIA_ID IN (" & filtro2 & ") "
'            Else
'                f2 = "   OR A.SUBFAMILIA_ID IN (" & filtro2 & ") "
'            End If
        End If
        If nadcap = True Then
            f3 = f3 & " AND A.NADCAP = 1"
        End If
        If mtl = True Then
            f3 = f3 & " AND A.MTL = 1"
        End If
        If EQA = True Then
            f3 = f3 & " AND A.EQA = 1"
        End If
        If ENAC = True Then
            f3 = f3 & " AND A.ENAC = 1"
        End If
        consulta = "SELECT A.NOMBRE,A.CODIGO,CASE WHEN A.NO_EDICION = 0 THEN A.EDICION ELSE 'N.A.' END,A.FECHA " & _
                   "  FROM ca_documentos A, decodificadora B " & _
                   " WHERE USO = 1 AND A.FAMILIA_ID = B.VALOR AND B.CODIGO = 8" & _
                   f1 & f2 & f3 & _
                   " AND A.ANULADO = 0 " & _
                   " ORDER BY B.PARAMETROS ,A.CODIGO,A.ID_DOCUMENTO"
        Set Listado_Auditoria = datos_bd(consulta)
End Function

'E0312-I
' Se muestran los documentos de calidad de la familia 'PNT C
Public Function Listado_Combo_procedimientos_calibracion() As ADODB.Recordset
        Dim consulta As String

        consulta = "SELECT A.ID_DOCUMENTO, A.NOMBRE " & _
                   "  FROM ca_documentos A, decodificadora B, decodificadora C " & _
                   " WHERE (A.FAMILIA_ID = B.VALOR AND B.CODIGO = 8) " & _
                   "   AND (A.ESTADO_ID = C.VALOR AND C.CODIGO = 7) " & _
                   "   AND A.NOMBRE LIKE '%%' " & _
                   "   AND A.CODIGO LIKE '%%' " & _
                   "   AND A.FAMILIA_ID = 14 " & _
                   "   AND A.USO = 1 " & _
                   "   AND A.ANULADO = 0 " & _
                   " ORDER BY A.ID_DOCUMENTO "
                   
        Set Listado_Combo_procedimientos_calibracion = datos_bd(consulta)
End Function

Public Function Listado_POR_NORMA(NORMA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_DOCUMENTO, A.NOMBRE,A.EDICION, A.FECHA,C.DESCRIPCION " & _
                    " FROM ca_documentos A, CA_DOCUMENTOS_NORMAS B, decodificadora C " & _
                    " where b.NORMA_ID = " & NORMA & _
                    "   AND A.ID_DOCUMENTO = B.DOCUMENTO_ID " & _
                    "   AND C.CODIGO = " & DECODIFICADORA.CA_DOCUMENTOS_RESPONSABLES & _
                    "  AND C.VALOR = A.RESPONSABLE_ID " & _
                    " order by A.CODIGO"
        Set Listado_POR_NORMA = datos_bd(consulta)
End Function

Public Function Listado_Combo_procedimientos_verificacion() As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT A.ID_DOCUMENTO, A.NOMBRE " & _
                   "  FROM ca_documentos A, decodificadora B, decodificadora C " & _
                   " WHERE (A.FAMILIA_ID = B.VALOR AND B.CODIGO = 8) " & _
                   "   AND (A.ESTADO_ID = C.VALOR AND C.CODIGO = 7) " & _
                   "   AND A.NOMBRE LIKE '%%' " & _
                   "   AND A.CODIGO LIKE '%%' " & _
                   "   AND A.FAMILIA_ID = 14 " & _
                   "   AND A.USO = 1 " & _
                   "   AND A.ANULADO = 0 " & _
                   " ORDER BY A.ID_DOCUMENTO "
        
        Set Listado_Combo_procedimientos_verificacion = datos_bd(consulta)
End Function
Public Function Informar_plan_formacion(ID As Long, PLAN As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ca_documentos SET " & _
                        " PLAN_ID = " & PLAN & _
                        " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        Informar_plan_formacion = True
        Exit Function
fallo:
        Informar_plan_formacion = False
        MsgBox "Error al Informar_plan_formacion (clsCa_documentos)", vbCritical, Err.Description
End Function

Public Function Modificar_Edicion(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ca_documentos SET " & _
                        " EDICION = '" & EDICION & "'," & _
                        " FECHA = '" & fecha & "'" & _
                        " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        Modificar_Edicion = True
        Exit Function
fallo:
        Modificar_Edicion = False
        MsgBox "Error al Modificar_Edicion (clsCa_documentos)", vbCritical, Err.Description
End Function
Public Function Modificar_Estado(ID As Long, ESTADO As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ca_documentos SET " & _
                        " ESTADO_ID = " & ESTADO & _
                        " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        Modificar_Estado = True
        Exit Function
fallo:
        Modificar_Estado = False
        MsgBox "Error al Modificar_Estado (clsCa_documentos)", vbCritical, Err.Description
End Function

Public Function Verificar_Codigo(CODIGO As String, documento As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ca_documentos WHERE ANULADO=0 AND ucase(CODIGO) = '" & UCase(CODIGO) & "' AND ID_DOCUMENTO <> " & documento
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Verificar_Codigo = False
        Else
            Verificar_Codigo = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Verificar_Codigo = False
End Function
Public Function mostrar(ID As Long, MODAL As Boolean) As Boolean
    Dim oCA_Documento As New clsCa_documentos
    Dim oDOCUMENTO As New clsDocumentacion
   On Error GoTo Mostrar_Error
    Dim r As Long
    oCA_Documento.Carga (ID)
'    If oCA_Documento.getRUTA = "" Then
'        MsgBox "El documento " & oCA_Documento.getCODIGO & " no tiene asignado ningún documento.", vbExclamation, App.Title
'        Exit Function
'    End If
    
    Dim oImpresion As New clsImpresion
    If oImpresion.cargar_documento(ID, TIPOS_DOCUMENTOS_FIRMA.CA_DOCUMENTO) = True Then
        MsgBox "El documento se esta generando aún en el servidor. Espere unos segundos y vuelva e intentarlo.", vbInformation, App.Title
        Exit Function
    End If
    Set oImpresion = Nothing
    Dim destino As String
    
    If oCA_Documento.getDOCUMENTO_VINCULADO = 1 Then
        destino = oDOCUMENTO.CargarDocumento(TOBJETO.TOBJETO_CA_DOCUMENTO, ID, 0, True, True)
        If destino = "" Then
            MsgBox "No encuentro el documento asociado.", vbCritical, App.Title
        End If
    Else
        destino = Replace(oCA_Documento.getRUTA, "/", "\")
    '    If Dir(destino) <> "" Then
            If UCase(Right(destino, 3)) = "XLS" Or UCase(Right(destino, 4)) = "XLSX" Or UCase(Right(destino, 4)) = "XLSM" Then
                If Dir(destino) <> "" Then
                    Dim XLA As excel.Application
                    Dim XLW As excel.Workbook
                    Dim XLS As excel.Worksheet
                    Set XLA = New excel.Application
                    Set XLW = XLA.Workbooks.Open(destino, , True)
                    Set XLS = XLW.Worksheets(1)
                    XLA.visible = True
                Else
                    MsgBox "No encuentro el documento asociado.", vbCritical, App.Title
                End If
            Else
                Dim oParametro As New clsParametros
                oParametro.Carga parametros.AcroPDF, ""
    '            Dim fso As New FileSystemObject
    '            Dim temp As String
    '            If fso.FileExists(destino) Then
    '                Dim cad() As String
    '                cad = Split(destino, "\")
    '                temp = DIRECTORIO_TEMPORAL & "\" & cad(UBound(cad))
    '                fso.CopyFile destino, temp, True
    '                destino = temp
    '            End If
    '            Set fso = Nothing
                destino = oDOCUMENTO.CargarDocumento(TOBJETO.TOBJETO_CA_DOCUMENTO, ID, 0, True, False)
                If oParametro.getVALOR = 1 Then
                    If MODAL Then
                        frmPrevisualizarPDF.nombre = oCA_Documento.getNOMBRE
                        frmPrevisualizarPDF.PK = ID
                        frmPrevisualizarPDF.tipo = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_DOCUMENTO
                        frmPrevisualizarPDF.RUTA = destino
                        frmPrevisualizarPDF.Show 1
                    Else
                        Dim oform As New frmPrevisualizarPDF
                        oform.nombre = oCA_Documento.getNOMBRE
                        oform.PK = ID
                        oform.tipo = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_DOCUMENTO
                        oform.RUTA = destino
                        oform.Show
                        Set oform = Nothing
                    End If
                Else
                    r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
                End If
            End If
            ' Si es la primera vez que el usuario abre el documento, inserta un registro en la tabla de firmas
            ' para que lo valore
    'M2394-I
    '        Dim oCDV As New clsCa_documentos_val
    '        oCDV.verificarExisteEvaluacion ID
    '        Set oCDV = Nothing
    'M2394-F
    '    Else
    '        MsgBox "No encuentro el documento asociado.", vbCritical, App.Title
    '    End If
    End If
            


   On Error GoTo 0
   Exit Function

Mostrar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Mostrar of Módulo de clase clsCa_documentos"
End Function
Public Function mostrarEdicion(ID As Long, EDICION As Integer, MODAL As Boolean) As Boolean
    Dim oCA_Documento As New clsCa_documentos
    Dim oDOCUMENTO As New clsDocumentacion
    oCA_Documento.Carga ID
   On Error GoTo Mostrar_Error
    Dim r As Long
    Dim destino As String
    oCA_Documento.Carga ID
    If oCA_Documento.getDOCUMENTO_VINCULADO = 1 Then
        destino = oDOCUMENTO.CargarDocumento(TOBJETO.TOBJETO_CA_DOCUMENTO, ID, EDICION, True, False)
    Else
        If oCA_Documento.getPLANTILLA_ID = 2 Or oCA_Documento.getPLANTILLA_ID = 5 Then
'            If oca_documento.getEDICION = EDICION Then
'            If oca_documento.getEDICION = EDICION Then
            If oCA_Documento.getEDICION = EDICION And (oCA_Documento.getESTADO_ID = C_CA_DOCUMENTOS_ESTADOS.CA_VIGOR Or oCA_Documento.getESTADO_ID = C_CA_DOCUMENTOS_ESTADOS.CA_CADUCADO) Then
                destino = Replace(oCA_Documento.getRUTA, "/", "\")
            Else
                'Buscar en el histórico
                Dim EXTENSION As String
                Dim RUTA_VERSIONES As String
                Dim oDeco As New clsDecodificadora
                oDeco.Carga_valor DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS, oCA_Documento.getFAMILIA_ID
                RUTA_VERSIONES = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\Trabajo\Versiones\"
                RUTA_VERSIONES = RUTA_VERSIONES & oDeco.getDESCRIPCION
                ' Extension
                oDeco.Carga_valor DECODIFICADORA.CALIDAD_PLANTILLAS_DOCUMENTOS, oCA_Documento.getPLANTILLA_ID
                Dim s() As String
                s = Split(oDeco.getPARAMETROS, ".")
                EXTENSION = "." & s(1)
                ' Nombre del documento y su copia (version)
                destino = RUTA_VERSIONES & "\" & Replace(Eliminar_Caracteres_Archivo(Trim(oCA_Documento.getCODIGO)), ".", " ") & " Ed." & EDICION & EXTENSION
    '            destino = Replace(oca_documento.getRUTA, "/", "\")
            End If
        Else
            If oCA_Documento.getEDICION = EDICION Then
                destino = oDOCUMENTO.CargarDocumento(TOBJETO.TOBJETO_CA_DOCUMENTO, ID, EDICION, True, False)
            Else
                destino = oDOCUMENTO.CargarDocumento(TOBJETO.TOBJETO_CA_DOCUMENTO, ID, EDICION, False, False)
            End If
        End If
    End If
    If destino = "" Then
        MsgBox "El documento " & oCA_Documento.getCODIGO & " no tiene asignado ningún documento.", vbExclamation, App.Title
        Exit Function
    End If
    If Dir(destino) <> "" Then
        If UCase(Right(destino, 3)) = "XLS" Then
            Dim XLA As excel.Application
            Dim XLW As excel.Workbook
            Dim XLS As excel.Worksheet
            Set XLA = New excel.Application
            Set XLW = XLA.Workbooks.Open(destino, , True)
            Set XLS = XLW.Worksheets(1)
            XLA.visible = True
        Else
            Dim oParametro As New clsParametros
            oParametro.Carga parametros.AcroPDF, ""
            If oParametro.getVALOR = 1 Then
                If MODAL Then
                    frmPrevisualizarPDF.nombre = oCA_Documento.getNOMBRE
                    frmPrevisualizarPDF.PK = ID
                    frmPrevisualizarPDF.tipo = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_DOCUMENTO
                    frmPrevisualizarPDF.RUTA = destino
                    frmPrevisualizarPDF.Show 1
                Else
                    Dim oform As New frmPrevisualizarPDF
                    oform.nombre = oCA_Documento.getNOMBRE
                    oform.PK = ID
                    oform.tipo = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_DOCUMENTO
                    oform.RUTA = destino
                    oform.Show
                    Set oform = Nothing
                End If
            Else
                r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
            End If
        End If
    Else
        MsgBox "No encuentro el documento asociado.", vbCritical, App.Title
    End If
   On Error GoTo 0
   Exit Function

Mostrar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Mostrar of Módulo de clase clsCa_documentos"
End Function

Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CA_DOCUMENTOS"
        .setDESCRIPCION = "Documentos"
        .setPK = "ID_DOCUMENTO"
        .setCAMPO = "NOMBRE"
        .setFILTRO = filtro
'BUG-XXXX-I
'       .setMUESTRA_DETALLE = True
        If FK <> 0 Then
           .setMUESTRA_DETALLE = True
           Set .FORMULARIO = FORMULARIO
        Else
           .setMUESTRA_DETALLE = False
        End If
'BUG-XXXX-F
    End With
End Function

Public Function esPNT(ID As Long) As Boolean
   On Error GoTo esPNT_Error

        Me.Carga ID
        If Me.getPLANTILLA_ID = 1 Or Me.getPLANTILLA_ID = 3 Then
            esPNT = True
        Else
            esPNT = False
        End If

   On Error GoTo 0
   Exit Function

esPNT_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure esPNT of Módulo de clase clsCa_documentos"
End Function


