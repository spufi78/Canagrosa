VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_NORMAS
'***************************************************************
Private ID_NORMA As Long
Private nombre As String
Private CODIGO As String
Private EDICION As String
Private fecha As String
Private FECHA_REVISION As String
Private ESTADO_ID As Integer
Private TIPO_ID As Integer
Private SECTOR_ID As Integer
Private AFECTA As Integer
Private nadcap As Integer
Private mtl As Integer
Private ENAC As Integer
Private EQA As Integer
Private REVISION As Integer
Private OBSERVACIONES As String
Private RESPONSABLE_ID As Integer
Private ruta As String
Private FTP As Integer
Private uso As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_NORMAS
'***************************************************************
Public Property Let setID_NORMA(ByVal dato As Long)
        ID_NORMA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setEDICION(ByVal dato As String)
        EDICION = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setFECHA_REVISION(ByVal dato As String)
        FECHA_REVISION = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Integer)
        ESTADO_ID = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Integer)
        TIPO_ID = dato
End Property
Public Property Let setSECTOR_ID(ByVal dato As Integer)
        SECTOR_ID = dato
End Property
Public Property Let setAFECTA(ByVal dato As Integer)
        AFECTA = dato
End Property
Public Property Let setNADCAP(ByVal dato As Integer)
        nadcap = dato
End Property
Public Property Let setMTL(ByVal dato As Integer)
        mtl = dato
End Property
Public Property Let setENAC(ByVal dato As Integer)
        ENAC = dato
End Property
Public Property Let setEQA(ByVal dato As Integer)
        EQA = dato
End Property
Public Property Let setREVISION(ByVal dato As Integer)
        REVISION = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Integer)
        RESPONSABLE_ID = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        ruta = dato
End Property
Public Property Let setFTP(ByVal dato As Integer)
        FTP = dato
End Property
Public Property Let setUSO(ByVal dato As Integer)
        uso = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_NORMAS
'***************************************************************
Public Property Get getID_NORMA() As Long
        getID_NORMA = ID_NORMA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getEDICION() As String
        getEDICION = EDICION
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getFECHA_REVISION() As String
        getFECHA_REVISION = FECHA_REVISION
End Property
Public Property Get getESTADO_ID() As Integer
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getTIPO_ID() As Integer
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getSECTOR_ID() As Integer
        getSECTOR_ID = SECTOR_ID
End Property
Public Property Get getAFECTA() As Integer
        getAFECTA = AFECTA
End Property
Public Property Get getNADCAP() As Integer
        getNADCAP = nadcap
End Property
Public Property Get getMTL() As Integer
        getMTL = mtl
End Property
Public Property Get getENAC() As Integer
        getENAC = ENAC
End Property
Public Property Get getEQA() As Integer
        getEQA = EQA
End Property
Public Property Get getREVISION() As Integer
        getREVISION = REVISION
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getRESPONSABLE_ID() As Integer
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getRUTA() As String
        getRUTA = ruta
End Property
Public Property Get getFTP() As Integer
        getFTP = FTP
End Property
Public Property Get getUSO() As Integer
        getUSO = uso
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_NORMAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_NORMAS WHERE ID_NORMA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_NORMA = rs("ID_NORMA")
                nombre = rs("NOMBRE")
                CODIGO = rs("CODIGO")
                EDICION = rs("EDICION")
                fecha = rs("FECHA")
                FECHA_REVISION = rs("FECHA_REVISION")
                ESTADO_ID = rs("ESTADO_ID")
                TIPO_ID = rs("TIPO_ID")
                SECTOR_ID = rs("SECTOR_ID")
                AFECTA = rs("AFECTA")
                nadcap = rs("NADCAP")
                mtl = rs("MTL")
                ENAC = rs("ENAC")
                EQA = rs("EQA")
                REVISION = rs("REVISION")
                OBSERVACIONES = rs("OBSERVACIONES")
                RESPONSABLE_ID = rs("RESPONSABLE_ID")
                ruta = rs("RUTA")
                FTP = rs("FTP")
                uso = rs("USO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_normas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_NORMA) FROM CA_NORMAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_NORMA = 1
        Else
                ID_NORMA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO CA_NORMAS " & _
                   "  VALUES (" & _
                        ID_NORMA & "," & "'" & nombre & "'" & "," & "'" & CODIGO & "'" & "," & _
                        "'" & EDICION & "'" & "," & "'" & fecha & "'" & "," & "'" & FECHA_REVISION & "'" & "," & _
                        ESTADO_ID & "," & TIPO_ID & "," & SECTOR_ID & "," & _
                        AFECTA & "," & nadcap & "," & mtl & "," & ENAC & "," & _
                        EQA & "," & REVISION & "," & "'" & OBSERVACIONES & "'" & "," & RESPONSABLE_ID & "," & _
                        "'" & ruta & "'" & "," & FTP & "," & uso & _
                ")"
        execute_bd consulta
        Insertar = ID_NORMA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_normas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_NORMAS SET " & _
                        " NOMBRE = '" & nombre & "'," & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " EDICION = '" & EDICION & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " FECHA_REVISION = '" & FECHA_REVISION & "'," & _
                        " ESTADO_ID = " & ESTADO_ID & "," & _
                        " TIPO_ID = " & TIPO_ID & "," & _
                        " SECTOR_ID = " & SECTOR_ID & "," & _
                        " AFECTA = " & AFECTA & "," & _
                        " NADCAP = " & nadcap & "," & _
                        " MTL = " & mtl & "," & _
                        " ENAC = " & ENAC & "," & _
                        " EQA = " & EQA & "," & _
                        " REVISION = " & REVISION & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " RESPONSABLE_ID = " & RESPONSABLE_ID & "," & _
                        " FTP = " & FTP & "," & _
                        " USO = " & uso & "" & _
                " WHERE ID_NORMA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_normas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        'E0040-I
        ' Antes de borrar una norma hay que comprobar si algún equipo la tiene vinculada
        Dim rs As ADODB.Recordset
        
        'consulta = "SELECT * FROM EQUIPOS_NORMAS WHERE NORMA_ID = " & id
        
        ' Jonathan.2010.08 -> Los equipos pueden tener vinculados documentos de calidad.
        
        'consulta = "SELECT * FROM eq_normas_equipos WHERE documento_ID = " & id
        'Set rs = datos_bd(consulta)
        'If rs.RecordCount <> 0 Then
        '    MsgBox "No se puede eliminar. Existen equipos con esta norma.", vbCritical, App.Title
        '    Eliminar = False
        '    Exit Function
        'End If
        'E0040-F
        consulta = "DELETE FROM CA_NORMAS " & _
                "    WHERE ID_NORMA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_normas)", vbCritical, Err.Description
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_NORMA,NOMBRE FROM CA_NORMAS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado(tipo As String, SECTOR As String, ESTADO As String, nombre As String, CODIGO As String, subtipo As String, ENAC As Integer, nadcap As Integer, mtl As Integer, EQA As Integer, FTP As Integer, estado_consulta As Integer) As ADODB.Recordset
        Dim consulta As String
        Dim oTIPO As String
        Dim oSECTOR As String
        Dim oestado As String
        Dim oENAC As String
        Dim oNADCAP As String
        Dim oMTL As String
        Dim oEQA As String
        Dim osubtipo As String
        Dim oFtp As String
        Dim oconsulta As String
        oTIPO = ""
        oestado = ""
        osubtipo = ""
        If tipo <> "0" Then
            oTIPO = " AND A.TIPO_ID = " & CInt(tipo)
        End If
        If SECTOR <> "0" Then
            oSECTOR = " AND A.SECTOR_ID = " & CInt(SECTOR)
        End If
        If ESTADO <> "0" Then
            oestado = " AND A.ESTADO_ID = " & CInt(ESTADO)
        End If
        If subtipo <> "0" Then
            osubtipo = " AND A.AFECTA = " & CInt(subtipo)
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
        If FTP = 1 Then
            oFtp = " AND A.FTP = 1 "
        End If
        If estado_consulta = 1 Then
'            oconsulta = " AND A.ESTADO_ID <> 25 "
            oconsulta = " AND A.ESTADO_ID =  25 " ' Solo las no controladas, Marga el 11/03/2014
        Else
            oconsulta = " AND A.ESTADO_ID <> 25 "
        End If
        consulta = "SELECT A.ID_NORMA,A.NOMBRE,B.DESCRIPCION,D.DESCRIPCION,A.CODIGO,A.EDICION,A.FECHA,C.DESCRIPCION,A.FTP,A.FECHA_REVISION " & _
                   "  FROM CA_NORMAS A, decodificadora B,decodificadora C,decodificadora D " & _
                   " WHERE (A.TIPO_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.CA_NORMAS_TIPOS & ") " & _
                   "   AND (A.ESTADO_ID = C.VALOR AND C.CODIGO = " & DECODIFICADORA.CA_NORMAS_ESTADOS & ") " & _
                   "   AND (A.AFECTA = D.VALOR AND D.CODIGO = " & DECODIFICADORA.CA_NORMAS_SUBTIPOS & ") " & _
                   "   AND A.NOMBRE LIKE '%" & nombre & "%' " & _
                   "   AND A.CODIGO LIKE '%" & CODIGO & "%' " & _
                   oTIPO & _
                   oSECTOR & _
                   oestado & _
                   osubtipo & _
                   oENAC & oNADCAP & oMTL & oEQA & oFtp & oconsulta & _
                   " ORDER BY A.NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_Codigo(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_NORMA,A.NOMBRE,B.DESCRIPCION,D.DESCRIPCION,A.CODIGO,A.EDICION,A.FECHA,C.DESCRIPCION,A.FTP,A.FECHA_REVISION " & _
                   "  FROM CA_NORMAS A, decodificadora B, decodificadora C, decodificadora D " & _
                   " WHERE (A.TIPO_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.CA_NORMAS_TIPOS & ") " & _
                   "   AND (A.ESTADO_ID = C.VALOR AND C.CODIGO = " & DECODIFICADORA.CA_NORMAS_ESTADOS & ") " & _
                   "   AND (A.AFECTA = D.VALOR AND D.CODIGO = " & DECODIFICADORA.CA_NORMAS_SUBTIPOS & ") " & _
                   "   AND A.ID_NORMA = " & ID
        Set Listado_por_Codigo = datos_bd(consulta)
End Function
Public Function Informar_ruta(ID As Long, sruta As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_NORMAS SET " & _
                        " RUTA = '" & sruta & "'" & _
                " WHERE ID_NORMA = " & ID
        execute_bd consulta
        Informar_ruta = True
        Exit Function
fallo:
        Informar_ruta = False
        MsgBox "Error al informar la ruta de la norma (clsCa_Normas)", vbCritical, Err.Description
End Function


Public Function Listado_Auditoria(filtro1 As String, filtro2 As String, nadcap As Boolean, mtl As Boolean, EQA As Boolean, ENAC As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim f1 As String
        Dim f2 As String
        Dim f3 As String
        Dim f4 As String
        If filtro1 <> "" Then
            f1 = "   AND A.TIPO_ID IN (" & filtro1 & ") "
        End If
        If filtro2 <> "" Then
'            If f1 = "" Then
                f2 = "   AND A.AFECTA IN (" & filtro2 & ") "
'            Else
'                f2 = "   OR A.AFECTA IN (" & filtro2 & ") "
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
        ' Omitir las normas en los siguientes estados
        ' CONSULTA = 25 ' Rosario
        f4 = " AND A.ESTADO_ID <> 25"
        consulta = "SELECT A.NOMBRE,A.CODIGO,A.EDICION,A.FECHA,B.DESCRIPCION,A.NADCAP,A.EQA,A.ENAC " & _
                   "  FROM CA_NORMAS A, decodificadora B " & _
                   " WHERE (A.AFECTA = B.VALOR AND B.CODIGO = " & DECODIFICADORA.CA_NORMAS_SUBTIPOS & ") " & _
                   f1 & f2 & f3 & f4 & _
                   " AND A.ESTADO_ID <> 25 " & _
                   " ORDER BY A.TIPO_ID,A.AFECTA,A.NOMBRE"
        Set Listado_Auditoria = datos_bd(consulta)
End Function

Public Function mostrar(ID As Long, MODAL As Boolean) As Boolean
    Dim r As String
    Dim destino As String
    Dim oNorma As New clsCa_normas
    Dim oParametro As New clsParametros
    oNorma.Carga (ID)
    ' BD DOC
    oParametro.Carga parametros.BD_DOCUMENTACION_NORMAS, ""
    If oParametro.getVALOR = 1 Then
        Dim oDoc As New clsDocumentacion
        destino = oDoc.CargarDocumento(TOBJETO.TOBJETO_CA_NORMA, ID, 0, True, False)
        If destino = "" Then
            MsgBox "No existe el documento solicitado.", vbCritical, App.Title
            Exit Function
        End If
    Else
        If oNorma.getRUTA = "" Then
            MsgBox "La NORMA no tiene asignado documento.", vbExclamation, App.Title
            Exit Function
        End If
        destino = Replace(oNorma.getRUTA, "/", "\")
    End If
    On Error GoTo fallo
    If Dir(destino) <> "" Then
      If UCase(Right(destino, 3)) = "PDF" Then
        oParametro.Carga parametros.AcroPDF, ""
        If oParametro.getVALOR = 1 Then
        
            If MODAL Then
                frmPrevisualizarPDF.nombre = oNorma.getNOMBRE
                frmPrevisualizarPDF.PK = ID
                frmPrevisualizarPDF.tipo = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_NORMA
                frmPrevisualizarPDF.ruta = destino
                frmPrevisualizarPDF.Show 1
            Else
                Dim oform As New frmPrevisualizarPDF
                oform.nombre = oNorma.getNOMBRE
                oform.PK = ID
                oform.tipo = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_NORMA
                oform.ruta = destino
                oform.Show
                Set oform = Nothing
            End If
                
        Else
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
        End If
      Else
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
      End If
    Else
        MsgBox "No encuentro el documento en la ruta especifica.", vbCritical, App.Title
    End If
    Exit Function
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title & ":" & Err.Description
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    Dim consulta As String
    consulta = "SELECT ID_NORMA,CONCAT(NOMBRE,' (',CODIGO,')') " & _
               "  FROM CA_NORMAS " & _
               " WHERE 1=1 " & _
               filtro
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CA_NORMAS"
        .setDESCRIPCION = "Normas"
        .setPK = "ID_NORMA"
        .setCAMPO = "CONCAT(NOMBRE,' (',CODIGO,')')"
        .setFILTRO = filtro
        .setQUERY = consulta
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function

