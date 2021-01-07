VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRpr_botes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSRPR_BOTES
'***************************************************************
Private ID_BOTE_PR As Long
Private TIPO_REACTIVO_PR_ID As Long
Private NUMERO As Long
Private VOLUMEN As String
Private OBSERVACIONES As String
Private fecha_fabricacion As String
Private fecha_caducidad As String
Private fecha_fin As String
Private SEGUN_USO As Integer
Private EMPLEADO_ID As Long
Private FTIMESTAMP As String
'E0172-I
Private TIPO_ID As Long
'E0172-F
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSRPR_BOTES
'***************************************************************
Public Property Let setID_BOTE_PR(ByVal dato As Long)
        ID_BOTE_PR = dato
End Property
Public Property Let setTIPO_REACTIVO_PR_ID(ByVal dato As Long)
        TIPO_REACTIVO_PR_ID = dato
End Property
Public Property Let setNUMERO(ByVal dato As Long)
        NUMERO = dato
End Property
Public Property Let setVOLUMEN(ByVal dato As String)
        VOLUMEN = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setFECHA_FABRICACION(ByVal dato As String)
        fecha_fabricacion = dato
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
        fecha_caducidad = dato
End Property
Public Property Let setFECHA_FIN(ByVal dato As String)
        fecha_fin = dato
End Property
Public Property Let setSEGUN_USO(ByVal dato As Integer)
        SEGUN_USO = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setFTIMESTAMP(ByVal dato As String)
        FTIMESTAMP = dato
End Property
'E0173-I
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
'E0173-F
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSRPR_BOTES
'***************************************************************
Public Property Get getID_BOTE_PR() As Long
        getID_BOTE_PR = ID_BOTE_PR
End Property
Public Property Get getTIPO_REACTIVO_PR_ID() As Long
        getTIPO_REACTIVO_PR_ID = TIPO_REACTIVO_PR_ID
End Property
Public Property Get getNUMERO() As Long
        getNUMERO = NUMERO
End Property
Public Property Get getVOLUMEN() As String
        getVOLUMEN = VOLUMEN
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFECHA_FABRICACION() As String
        getFECHA_FABRICACION = fecha_fabricacion
End Property
Public Property Get getFECHA_CADUCIDAD() As String
        getFECHA_CADUCIDAD = fecha_caducidad
End Property
Public Property Get getFECHA_FIN() As String
        getFECHA_FIN = fecha_fin
End Property
Public Property Get getSEGUN_USO() As Integer
        getSEGUN_USO = SEGUN_USO
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getFTIMESTAMP() As String
        getFTIMESTAMP = FTIMESTAMP
End Property
'E0174-I
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
'E0174-F
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSRPR_BOTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM RPR_BOTES WHERE ID_BOTE_PR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_BOTE_PR = rs("ID_BOTE_PR")
                TIPO_REACTIVO_PR_ID = rs("TIPO_REACTIVO_PR_ID")
                NUMERO = rs("NUMERO")
                VOLUMEN = rs("VOLUMEN")
                If Not IsNull(rs("OBSERVACIONES")) Then
                    OBSERVACIONES = rs("OBSERVACIONES")
                Else
                    OBSERVACIONES = ""
                End If
                fecha_fabricacion = rs("FECHA_FABRICACION")
                fecha_caducidad = rs("FECHA_CADUCIDAD")
                If Not IsNull(rs("FECHA_FIN")) Then
                    fecha_fin = rs("FECHA_FIN")
                End If
                SEGUN_USO = rs("SEGUN_USO")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                FTIMESTAMP = rs("FTIMESTAMP")
                'E0174-I
                TIPO_ID = rs("TIPO_ID")
                'E0174-F
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRpr_botes)", vbCritical, Err.Description
End Function
Public Function CrearID(CODIGO As String)
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_BOTE_PR) FROM RPR_BOTES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_BOTE_PR = 1
        Else
                ID_BOTE_PR = rs.Fields(0) + 1
        End If
        consulta = "select max(numero) from rpr_botes a, rpr_tipos b " & _
                   " Where a.TIPO_REACTIVO_PR_ID = b.ID_TIPO_REACTIVO_PR " & _
                   "  and b.CODIGO = '" & CODIGO & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                NUMERO = 1
        Else
                NUMERO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rpr_tipo As New clsRPR_Tipos
        rpr_tipo.CARGAR TIPO_REACTIVO_PR_ID
        Me.CrearID rpr_tipo.getCODIGO
        'E0175-I
        consulta = "INSERT INTO RPR_BOTES " & _
                   "  VALUES (" & _
                        ID_BOTE_PR & "," & TIPO_REACTIVO_PR_ID & "," & NUMERO & "," & _
                        "'" & VOLUMEN & "','" & OBSERVACIONES & "'" & "," & "'" & fecha_fabricacion & "','" & fecha_caducidad & "',null," & SEGUN_USO & "," & _
                        USUARIO.getID_EMPLEADO & ",current_timestamp, " & TIPO_ID & ")"
        'E0175-F
        execute_bd consulta
        Insertar = ID_BOTE_PR
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRpr_botes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        'E0176-I
        consulta = "UPDATE RPR_BOTES SET " & _
                        " VOLUMEN = '" & VOLUMEN & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " FECHA_FABRICACION = '" & fecha_fabricacion & "'," & _
                        " FECHA_CADUCIDAD = '" & fecha_caducidad & "'," & _
                        " EMPLEADO_ID = " & USUARIO.getID_EMPLEADO & "," & _
                        " FTIMESTAMP = current_timestamp " & ", " & _
                        " SEGUN_USO = " & SEGUN_USO & ", " & _
                        " TIPO_ID = " & TIPO_ID & _
                " WHERE ID_BOTE_PR = " & ID
        'E0176-F
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRpr_botes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM RPR_BOTES " & _
                "    WHERE ID_BOTE_PR = " & ID
        execute_bd consulta
        consulta = "DELETE FROM RPR_BOTES_COMPONENTES " & _
                "    WHERE BOTE_PR_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRpr_botes)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM RPR_BOTES ORDER BY ID_BOTE_PR"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_tipo(tipo As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM RPR_BOTES A, RPR_TIPOS B " & _
                   " where A.TIPO_REACTIVO_PR_ID = " & tipo & _
                   "   AND A.TIPO_REACTIVO_PR_ID = B.ID_TIPO_REACTIVO_PR " & _
                   "   AND A.FECHA_FIN IS NULL " & _
                   " ORDER BY ID_BOTE_PR "
        Set Listado_por_tipo = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_BOTE_PR,TIPO_REACTIVO_PR_ID FROM RPR_BOTES ORDER BY TIPO_REACTIVO_PR_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    Dim consulta As String
    consulta = "SELECT A.ID_BOTE_PR,CONCAT(B.NOMBRE , ' (',B.CODIGO,'-',CAST(A.NUMERO AS CHAR),')') " & _
               "  FROM RPR_BOTES A, RPR_TIPOS B " & _
               " WHERE A.TIPO_REACTIVO_PR_ID = B.ID_TIPO_REACTIVO_PR " & _
               filtro
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "RPR_BOTES"
        .setDESCRIPCION = "Reactivos Propios"
        .setPK = "ID_BOTE_PR"
        .setCAMPO = "CONCAT(B.NOMBRE , ' (',B.CODIGO,'-',CAST(A.NUMERO AS CHAR),')')"
        .setQUERY = consulta
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function


Public Function imprimir_etiqueta(LISTA_REACTIVOS As String, Optional EQUIPO As String) As Boolean
    Dim prnPrinter As Printer
    
    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros

   On Error GoTo imprimir_etiqueta_Error
    log ("Impresion etiqueta reactivo")
    Dim PC As String
    If EQUIPO <> "" Then
        PC = EQUIPO
    Else
        PC = USUARIO.getUSO
    End If
    log ("PC : " & PC)
    Dim i As Integer
    If PC = "IBERIA" Then
        Dim oEtiqueta As New clsEtiquetas
        Dim Listar() As String
        Listar = Split(LISTA_REACTIVOS, ",")
        For i = LBound(Listar) To UBound(Listar)
            With oEtiqueta
                .setTIPO_ID = C_ETIQUETAS_TIPOS.C_ETIQUETAS_TIPOS_RPR
                .setCENTRO_ID = CENTROS.CENTRO_MADRID
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setID = Listar(i)
                .Insertar
            End With
        Next
        MsgBox "Se han enviado a imprimir las copias.", vbInformation, App.Title
        Exit Function
    End If
    
    If Not oParametro.Carga(parametros.IMPRESORA_ETIQUETAS_MEDIANA, PC) Then
        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
        Exit Function
    End If
'    Firmas.copiar_firma_usuario_activo
    log ("Parametro impresora : " & Replace(oParametro.getVALOR, "/", "\"))
    
    Dim impresora_encontrada As String
    Dim IMPRESORA As Printer
    impresora_encontrada = ""
    For Each prnPrinter In Printers
'        log "Impresora instalada : " & prnPrinter.DeviceName
        ' Quitamos del nombre de la impresora el #00X para las impresoras TSN del remoto
        Dim impresoraArray() As String
        Dim impresoraNombre As String
        impresoraArray = Split(UCase(prnPrinter.DeviceName), "#")
        impresoraNombre = Trim(impresoraArray(0))
        
'        If UCase(prnPrinter.DeviceName) = Replace(UCase(oParametro.getVALOR), "/", "\") Then
        If impresoraNombre = Replace(UCase(oParametro.getVALOR), "/", "\") Then
            Set Printer = prnPrinter
            Set IMPRESORA = prnPrinter
            impresora_encontrada = prnPrinter.DeviceName
            Exit For
        End If
    Next
    If impresora_encontrada <> "" Then
        Dim imp_ant As String
        Dim P1() As String
        Dim P2() As String
        ReDim P1(1) As String
        ReDim P2(1) As String
        P1(1) = "FIRMA"
        If ReadINI(App.Path + "\config.ini", "Documentos", "firmas") = "c:\geslab\recursos\firmas" Then
            Dim firma As String
            Dim lista() As String
            If USUARIO.getFIRMA <> "" Then
                lista = Split(USUARIO.getFIRMA, "/")
                firma = ReadINI(App.Path + "\config.ini", "Documentos", "firmas") & "\" & lista(UBound(lista))
            End If
            P2(1) = firma
        Else
            P2(1) = Replace(USUARIO.getFIRMA, "/", "\")
        End If
        imp_ant = Impresora_Predeterminada
        Establecer_Impresora impresora_encontrada
        With frmReport
            .iniciar
            .informe = "\RPR\rptRPR_Etiqueta_Pequeña"
'            .informe = "\REX\rptREX_Etiqueta"
            .criterio = "{rpr_botes.ID_BOTE_PR} in [" & LISTA_REACTIVOS & "]"
            .ParametrosNombre = P1
            .ParametrosValores = P2
            .imprimir = True
'            .imprimir = False
            .generar
'            .Show 1
        End With
'        Unload frmReport
        Set frmReport = Nothing
        Establecer_Impresora imp_ant
    Else
        MsgBox "No se localiza la impresora de etiquetas.", vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Function

imprimir_etiqueta_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Imprimir_Etiqueta of Módulo de clase clsBotes_ex"
End Function
Public Function Terminar(ID As Long, fecha As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE RPR_BOTES SET " & _
               " FECHA_FIN = '" & Format(fecha, "yyyy-mm-dd") & "'" & _
               " WHERE ID_BOTE_PR = " & ID
    execute_bd consulta
    Terminar = True
    
    Exit Function
fallo:
    Terminar = False
    MsgBox "Error al Terminar el Bote (Terminar)", vbCritical, Err.Description
End Function
'M1067-I
Public Function TerminarAnular(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE RPR_BOTES SET " & _
               " FECHA_FIN = NULL " & _
               " WHERE ID_BOTE_PR = " & ID
    execute_bd consulta
    TerminarAnular = True
    
    Exit Function
fallo:
    TerminarAnular = False
    MsgBox "Error al TerminarAnular el Bote (TerminarAnular)", vbCritical, Err.Description
End Function
'M1067-I

