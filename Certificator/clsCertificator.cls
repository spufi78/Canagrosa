VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCertificator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCERTIFICATOR
'***************************************************************
Private ID As Long
Private TOBJETO As Long
Private COBJETO As Long
Private EMPLEADO_ID As Long
Private TS As String
Private PC As String
Private ESTADO_ID As Long

Private EQUIPO_ID As Long
Private EQUIPO_NOMBRE As String
Private EQUIPO_NUMERO As String
Private EQUIPO_TIPO_ID As Long
Private EQUIPO_TIPO_EQUIPO As String
Private PLANTILLA As String
Private FILE As String


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCERTIFICATOR
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setTOBJETO(ByVal dato As Long)
        TOBJETO = dato
End Property
Public Property Let setCOBJETO(ByVal dato As Long)
        COBJETO = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
Public Property Let setPC(ByVal dato As String)
        PC = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Long)
        ESTADO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCERTIFICATOR
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getTOBJETO() As Long
        getTOBJETO = TOBJETO
End Property
Public Property Get getCOBJETO() As Long
        getCOBJETO = COBJETO
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
Public Property Get getPC() As String
        getPC = PC
End Property
Public Property Get getESTADO_ID() As Long
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getEQUIPO_TIPO_ID() As Long
        getEQUIPO_TIPO_ID = EQUIPO_TIPO_ID
End Property
Public Property Get getEQUIPO_TIPO_EQUIPO() As String
        getEQUIPO_TIPO_EQUIPO = EQUIPO_TIPO_EQUIPO
End Property

Public Property Get getEQUIPO_NOMBRE() As String
        getEQUIPO_NOMBRE = EQUIPO_NOMBRE
End Property
Public Property Get getEQUIPO_NUMERO() As String
        getEQUIPO_NUMERO = EQUIPO_NUMERO
End Property
Public Property Get getFILE() As String
        getFILE = FILE
End Property
Public Property Get getPLANTILLA() As String
        getPLANTILLA = PLANTILLA
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCERTIFICATOR
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "select a.ID,a.TOBJETO,a.COBJETO,a.EMPLEADO_ID,a.TS,a.PC,a.ESTADO_ID " & _
                    "     ,d.ID_EQUIPO as EQUIPO_ID,d.NOMBRE as EQUIPO_NOMBRE,d.NUMERO_EQUIPO_CLIENTE as EQUIPO_NUMERO " & _
                    "     ,CONCAT(b.NOMBRE,' ',b.APELLIDOS) AS NOMBRE,IFNULL(e.DESCRIPCION,''),IFNULL(e.FILE_NAME,'') AS PLANTILLA,a.TS,a.PC,a.ESTADO_ID,d.TIPO_EQUIPO_ID as EQUIPO_TIPO_ID, e.FILE,deco.DESCRIPCION as TIPO_EQUIPO " & _
                    " from certificator a  " & _
                    " left join geslab_canagrosa.usuarios b on a.EMPLEADO_ID = b.ID_EMPLEADO " & _
                    " left join geslab_canagrosa.eq_calibracion_equipos c on a.COBJETO = c.ID_CALIBRACION " & _
                    " left join geslab_canagrosa.equipos d on c.EQUIPO_ID = d.ID_EQUIPO " & _
                    " left join geslab_canagrosa_documentacion.equipos_plantillas e on d.TIPO_EQUIPO_ID = e.TIPO_EQUIPO_ID and vigor = 1 " & _
                    " left join geslab_canagrosa.decodificadora deco on deco.CODIGO = " & C_TIPOS_EQUIPO & " and deco.VALOR = d.TIPO_EQUIPO_ID " & _
                    " where a.ID = " & ID
        Set rs = datos_bd_metrologia(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                TS = rs("TS")
                PC = rs("PC")
                ESTADO_ID = rs("ESTADO_ID")
                PLANTILLA = rs("PLANTILLA")
                
                EQUIPO_ID = rs("EQUIPO_ID")
                EQUIPO_NOMBRE = rs("EQUIPO_NOMBRE")
                EQUIPO_NUMERO = rs("EQUIPO_NUMERO")
                EQUIPO_TIPO_ID = rs("EQUIPO_TIPO_ID")
                EQUIPO_TIPO_EQUIPO = rs("TIPO_EQUIPO")
                If Not IsNull(rs("FILE")) Then
                    FILE = rs("FILE")
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCertificator)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM CERTIFICATOR"
        Set rs = datos_bd_metrologia(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO CERTIFICATOR " & _
                   "  VALUES (" & _
                        ID & "," & TOBJETO & "," & COBJETO & "," & _
                        EMPLEADO_ID & "," & "'" & TS & "'" & "," & "'" & PC & "'" & "," & _
                        ESTADO_ID & _
                ")"
        execute_bd_metrologia consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCertificator)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CERTIFICATOR SET " & _
                        " TOBJETO = " & TOBJETO & "," & _
                        " COBJETO = " & COBJETO & "," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " TS = '" & TS & "'," & _
                        " PC = '" & PC & "'," & _
                        " ESTADO_ID = " & ESTADO_ID & "" & _
                " WHERE ID = " & ID
        execute_bd_metrologia consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCertificator)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Me.Carga ID
        consulta = "DELETE FROM CERTIFICATOR WHERE ID = " & ID
        execute_bd_metrologia consulta
        consulta = "DELETE FROM CERTIFICATOR_VIDA WHERE TOBJETO = " & Me.getTOBJETO & " AND COBJETO = " & Me.getCOBJETO
        execute_bd_metrologia consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCertificator)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long, TIPO_EQUIPO As Integer, EMPLEADO_ID As Integer, ESTADO_ID As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If ID <> 0 Then
            filtro = filtro & " AND a.ID = " & ID
        Else
            If TIPO_EQUIPO <> 0 Then
                filtro = filtro & " AND d.TIPO_EQUIPO_ID = " & TIPO_EQUIPO
            End If
            If EMPLEADO_ID <> 0 Then
                filtro = filtro & " AND a.EMPLEADO_ID = " & EMPLEADO_ID
            End If
            If ESTADO_ID <> "" Then
                filtro = filtro & " AND a.ESTADO_ID IN (" & ESTADO_ID & ")"
            End If
        End If
        consulta = "select a.ID,a.TOBJETO,a.COBJETO,d.ID_EQUIPO,d.NOMBRE,d.NUMERO_EQUIPO_CLIENTE,CONCAT(b.NOMBRE,' ',b.APELLIDOS) AS NOMBRE,IFNULL(e.DESCRIPCION,''),IFNULL(e.FILE_NAME,''),a.TS,a.PC,a.ESTADO_ID,d.TIPO_EQUIPO_ID " & _
                    " from certificator a  " & _
                    " left join geslab_canagrosa.usuarios b on a.EMPLEADO_ID = b.ID_EMPLEADO " & _
                    " left join geslab_canagrosa.eq_calibracion_equipos c on a.COBJETO = c.ID_CALIBRACION " & _
                    " left join geslab_canagrosa.equipos d on c.EQUIPO_ID = d.ID_EQUIPO " & _
                    " left join geslab_canagrosa_documentacion.equipos_plantillas e on d.TIPO_EQUIPO_ID = e.TIPO_EQUIPO_ID and vigor = 1 " & _
                    " where 1 = 1 " & _
                    filtro
        Set Listado = datos_bd_metrologia(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,TOBJETO FROM CERTIFICATOR ORDER BY TOBJETO"
        Set Listado_Combo = datos_bd_metrologia(consulta)
End Function
Public Function bloquear(ID As Long) As Boolean
        Dim consulta As String
   On Error GoTo bloquear_Error

        consulta = "UPDATE CERTIFICATOR SET PC = '" & frmMain.Winsock1.LocalHostName & "' WHERE ID = " & ID
        execute_bd_metrologia consulta
        
        Me.modificarEstado ID, C_ESTADOS.C_ESTADO_BLOQUEADO
        bloquear = True

   On Error GoTo 0
   Exit Function

bloquear_Error:
    bloquear = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bloquear of Módulo de clase clsCertificator"
End Function
Public Function desbloquear(ID As Long) As Boolean
        Dim consulta As String
   On Error GoTo bloquear_Error

        consulta = "UPDATE CERTIFICATOR SET PC = '' WHERE ID = " & ID
        execute_bd_metrologia consulta
        
        Me.modificarEstado ID, C_ESTADOS.C_ESTADO_PENDIENTE
        
        desbloquear = True

   On Error GoTo 0
   Exit Function

bloquear_Error:
    desbloquear = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure desbloquear of Módulo de clase clsCertificator"
End Function
Public Function modificarEstado(ID As Long, ESTADO_ID As Integer) As Boolean
        Dim consulta As String
   On Error GoTo bloquear_Error

        consulta = "UPDATE CERTIFICATOR SET ESTADO_ID = " & ESTADO_ID & " WHERE ID = " & ID
        execute_bd_metrologia consulta
        ' Insertar Vida
        Dim oC As New clsCertificator
        Dim oCV As New clsCertificator_vida
        oC.Carga ID
        With oCV
            .setTOBJETO = oC.getTOBJETO
            .setCOBJETO = oC.getCOBJETO
            .setEMPLEADO_ID = oC.getEMPLEADO_ID
            .setESTADO_ID = ESTADO_ID
            .Insertar
        End With
        modificarEstado = True

   On Error GoTo 0
   Exit Function

bloquear_Error:
    modificarEstado = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure desbloquear of Módulo de clase clsCertificator"
End Function
Public Function abrirCertificado(ID As Long) As Boolean
        On Error GoTo fallo
        If Me.Carga(ID) Then
            ' Mostrar certificado
            Dim CERTIFICADO As String
            Dim RUTA As String
            Dim fichero As String
            RUTA = App.Path & "\" & RUTA_CERTIFICADOS & "\"
            CERTIFICADO = Me.nombreCertificado(Me)
            fichero = RUTA & CERTIFICADO
            If Dir(fichero) = "" Then
                ' No existe excel, lo genero desde la plantilla
                Dim oEP As New clsEquipos_plantillas
                fichero = oEP.GeneraCertificado(Me.getEQUIPO_TIPO_ID, CERTIFICADO, False)
            End If
            If Dir(fichero) <> "" Then
                oEP.cumplimentarExcel fichero, ID
            End If
'            Else
'                ' Existe, lo abro
'                Dim iret As Long
'                iret = ShellExecute(0, vbNullString, RUTA & CERTIFICADO, vbNullString, App.Path, 1)
'            End If
            ' Bloquear
            Me.bloquear ID
        End If
        abrirCertificado = True
        Exit Function
fallo:
        abrirCertificado = False
        MsgBox "Error al abrirCertificado (clsCertificator)", vbCritical, Err.Description
End Function
Public Function finalizar(ID As Long) As Boolean
        On Error GoTo fallo
        If Me.Carga(ID) Then
            ' Insertar excel en bd
            Dim CERTIFICADO As String
            Dim RUTA As String
            RUTA = App.Path & "\" & RUTA_CERTIFICADOS & "\"
            CERTIFICADO = Me.nombreCertificado(Me)
            If Dir(RUTA & CERTIFICADO) = "" Then
                MsgBox "Error, el CERTIFICADO : " & CERTIFICADO & ", NO EXISTE!!", vbCritical, App.Title
                Exit Function
            End If
            Dim oD As New clsDocumentacion
            Dim salida As String
            salida = oD.SubirEquipo(Me.getEQUIPO_ID, C_TIPO_CERTIFICADO.C_CALIBRACION, Me.getCOBJETO, C_SUBTIPO_CERTIFICADO.C_HOJA, RUTA & CERTIFICADO, CERTIFICADO, Me.getEMPLEADO_ID)
            If salida <> "" Then
                MsgBox "Se ha producido un error al subir la hoja de calibracion : " & salida, vbCritical, App.Title
                Exit Function
            Else
                Me.actualizarRutasCalibracion C_SUBTIPO_CERTIFICADO.C_HOJA, Me.getCOBJETO, CERTIFICADO
            End If
            ' Generar PDF
            CERTIFICADO = Me.generarCertificadoPDF(ID, False)
            If CERTIFICADO <> "" Then
                ' Insertar PDF en BD de equipos
                salida = oD.SubirEquipo(Me.getEQUIPO_ID, C_TIPO_CERTIFICADO.C_CALIBRACION, Me.getCOBJETO, C_SUBTIPO_CERTIFICADO.C_CERTIFICADO, CERTIFICADO, Me.nombreCertificadoPDF(Me), Me.getEMPLEADO_ID)
                If salida <> "" Then
                    MsgBox "Se ha producido un error al subir el certificado de calibracion : " & salida, vbCritical, App.Title
                    Exit Function
                End If
                Me.actualizarRutasCalibracion C_SUBTIPO_CERTIFICADO.C_CERTIFICADO, Me.getCOBJETO, Me.nombreCertificadoPDF(Me)
                ' Insertar PDF en la MUESTRA
                salida = oD.SubirInformeMuestra(Me.getCOBJETO)
                
            End If
            ' Finalizar
            Me.modificarEstado ID, C_ESTADOS.C_ESTADO_FINALIZADO
        End If
        finalizar = True
        Exit Function
fallo:
        finalizar = False
        MsgBox "Error al abrirCertificado (clsCertificator)", vbCritical, Err.Description
End Function
Public Function actualizarRutasCalibracion(TIPO As Integer, ID_CAL As Long, CERTIFICADO As String) As Boolean
        Dim c As String
        Dim s As String
   On Error GoTo actualizarRutasCalibracion_Error

        Select Case TIPO
            Case C_SUBTIPO_CERTIFICADO.C_HOJA
                s = "ruta_plantilla"
            Case C_SUBTIPO_CERTIFICADO.C_CERTIFICADO
                s = "ruta_certificado"
            Case C_SUBTIPO_CERTIFICADO.C_EVALUACION
                s = "ruta_evaluacion"
        End Select
        
        If s <> "" Then
            c = "UPDATE geslab_canagrosa.eq_calibracion_equipos set " & s & " = '" & CERTIFICADO & "' WHERE ID_CALIBRACION = " & ID_CAL
            execute_bd_metrologia c
        End If
        actualizarRutasCalibracion = True

   On Error GoTo 0
   Exit Function

actualizarRutasCalibracion_Error:
    actualizarRutasCalibracion = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizarRutasCalibracion of Módulo de clase clsCertificator"
End Function
Public Function nombreCertificado(c As clsCertificator) As String
        Dim CERTIFICADO As String
        nombreCertificado = "C-" & c.getCOBJETO & "-" & Year(Date) & ".xlsx"
End Function
Public Function nombreCertificadoPDF(c As clsCertificator) As String
        nombreCertificadoPDF = "C-" & c.getCOBJETO & "-" & Year(Date) & ".pdf"
End Function
Public Function generarCertificadoPDF(ID As Long, Abrir As Boolean) As String
   On Error GoTo generarCertificadoPDF_Error
        Dim pdf As String
    If Me.Carga(ID) Then
        ' Insertar excel en bd
        Dim CERTIFICADO As String
        Dim RUTA As String
        RUTA = App.Path & "\" & RUTA_CERTIFICADOS & "\"
        On Error Resume Next
        MkDir RUTA & "pdf"
        On Error GoTo generarCertificadoPDF_Error
        
        CERTIFICADO = Me.nombreCertificado(Me)
        If Dir(RUTA & CERTIFICADO) = "" Then
            MsgBox "Error, el CERTIFICADO : " & CERTIFICADO & ", NO EXISTE!!", vbCritical, App.Title
            Exit Function
        End If
        
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Open(RUTA & CERTIFICADO)
        Set XLS = XLW.Worksheets(1)
        
        ' RECUPERO EL NOMBRE DE LA PESTAÑA DE IMPRESION
        Dim cert As String
        cert = XLS.Cells(CERTIFICADO_IMPRESION_FILA, CERTIFICADO_IMPRESION_COL)
'        MsgBox "Impresion de : " & cert
        
        Dim i As Integer
        For i = 1 To XLW.Worksheets.Count
            If XLW.Worksheets(i).Name = cert Then
                Set XLS = XLW.Worksheets(i)
                Exit For
            End If
        Next
        pdf = RUTA & "pdf\" & Me.nombreCertificadoPDF(Me)
        XLS.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
                    pdf, Quality:=xlQualityStandard, _
                    IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        XLW.Save
        XLA.Visible = False
        XLA.Quit
        Set XLS = Nothing
        Set XLW = Nothing
        Set XLA = Nothing
    
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, pdf, vbNullString, App.Path, 1)
        End If
    
        generarCertificadoPDF = pdf
    End If
    
   On Error GoTo 0
   Exit Function

cumplimentarExcel_Error:
    pdf = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cumplimentarExcel of Módulo de clase clsEquipos_plantillas"

   On Error GoTo 0
   Exit Function

generarCertificadoPDF_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generarCertificadoPDF of Módulo de clase clsCertificator"
End Function

