VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrevisualizar 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14805
   Icon            =   "frmPrevisualizar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   14805
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   9090
      TabIndex        =   20
      Top             =   8595
      Width           =   5640
      Begin VB.CommandButton cmdRevisar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Revisar"
         Height          =   1005
         Left            =   3060
         Picture         =   "frmPrevisualizar.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ESC-Salir"
         Height          =   1005
         Left            =   4365
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdManual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informe Manual"
         Height          =   1005
         Left            =   1755
         Picture         =   "frmPrevisualizar.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdAnexo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anexo Verificaciones"
         Height          =   1005
         Left            =   450
         Picture         =   "frmPrevisualizar.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   135
         Width           =   1275
      End
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   7830
      TabIndex        =   19
      Top             =   9000
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   7740
      TabIndex        =   18
      Top             =   8640
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archivos Adjuntos a la Muestra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   10845
      TabIndex        =   16
      Top             =   7560
      Visible         =   0   'False
      Width           =   3795
      Begin MSComctlLib.ListView lista 
         Height          =   1065
         Left            =   90
         TabIndex        =   17
         Top             =   225
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   1879
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame frmEdiciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ediciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   4500
      TabIndex        =   14
      Top             =   8505
      Visible         =   0   'False
      Width           =   3165
      Begin MSComctlLib.ListView listaEdiciones 
         Height          =   1065
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1879
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdacrobat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Abrir Acrobat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   12915
      Picture         =   "frmPrevisualizar.frx":1D68
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Generar una nueva edición del informe"
      Top             =   5535
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Generar nueva edición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   12600
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   2130
      Begin VB.CommandButton cmdInforme 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   300
         Picture         =   "frmPrevisualizar.frx":1E7E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Generar una nueva edición del informe"
         Top             =   1620
         Width           =   1545
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Previsualizar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   300
         Picture         =   "frmPrevisualizar.frx":3178
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Previsualizar nueva edición del informe de ensayo"
         Top             =   660
         Width           =   1545
      End
      Begin VB.TextBox txtedicion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         TabIndex        =   1
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   2
         Top             =   315
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   90
      TabIndex        =   5
      Top             =   8505
      Width           =   4380
      Begin VB.CheckBox chkCabecera 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ocultar cabecera"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   1035
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkimpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Muestra impresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2340
         TabIndex        =   12
         Top             =   630
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.CheckBox chkcorreo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enviada por correo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2340
         TabIndex        =   11
         Top             =   855
         Value           =   1  'Checked
         Width           =   1965
      End
      Begin VB.CommandButton cmdmail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-mail"
         Height          =   870
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Enviar informe de la ultima edición generada por E-mail"
         Top             =   315
         Width           =   1050
      End
      Begin VB.TextBox txtediciongen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   705
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   870
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir ultima edición generada"
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
         Height          =   195
         Index           =   0
         Left            =   2610
         TabIndex        =   8
         Top             =   315
         Width           =   525
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8370
      Top             =   8910
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin AcroPDFLibCtl.AcroPDF pdf1 
      Height          =   7800
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   13650
      _cx             =   5080
      _cy             =   5080
   End
End
Attribute VB_Name = "frmPrevisualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public PK As Long
Dim m_cTT As New cTooltip
Private Sub chkcorreo_Click()
    Dim oMuestra As New clsMuestra
    If chkCorreo.Value = Checked Then
        oMuestra.informar_correo PK, USUARIO.getID_EMPLEADO
        chkCorreo.BackColor = &HC0FFFF
    Else
        oMuestra.informar_correo PK, 0
        chkCorreo.BackColor = &HFF&
    End If
    Set oMuestra = Nothing
End Sub

Private Sub cmdAnexo_Click()
    Dim objrep As New frmReport
   On Error GoTo cmdAnexo_Click_Error

    With objrep
        .iniciar
        Dim c As String
        c = ""
        c = c & "{muestras.ID_MUESTRA} = " & PK
        c = c & " and {ce_recepcion_equipos.EN_INFORME} = 1.00 "
        c = c & " and {muestras.FECHA_COMIENZO} <= {eq_verificacion_equipos.FECHA_ACTUAL} "
        c = c & " and {muestras.FECHA_FINALIZACION} >= {eq_verificacion_equipos.FECHA_ACTUAL} "
        .criterio = c
        .informe = "Informes\rptCE_Anexo"
        .imprimir = False
        .pdf = ""
        .generar
        .Show 1
    End With
    Unload objrep

   On Error GoTo 0
   Exit Sub

cmdAnexo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnexo_Click of Formulario frmPrevisualizar"
End Sub

Private Sub cmdmanual_Click()
   On Error GoTo cmdmanual_Click_Error

        On Error Resume Next
    cd.DialogTitle = "Seleccione el pdf del informe"
    cd.Filter = "Archivos PDF de informes|*.pdf"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileTitle ' Solo título
        datos(1).Text = cd.FileName ' Completo
        If MsgBox("Va a insertar manualmente el informe de la muestra. ¿Desea Continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
        Dim oMuestra As New clsMuestra
        oMuestra.modificar_edicion_impresa PK, 1
        
        destino = NOMBRE_DOCUMENTO(PK, True) & ".pdf"
        
        Me.MousePointer = 11
        
        FileCopy datos(1).Text, destino

        If Dir(destino) <> "" Then
            'M1054-I
            oMuestra.modificarInformeManual PK, 1
            'M1054-F
            imprimir PK, 70, False
            MsgBox "Informe adjuntado correctamente.", vbOKOnly + vbInformation, App.Title
            mostrar_pdf PK, 0
        Else
            MsgBox "El informe no se ha adjuntado correctamente.", vbExclamation, App.Title
        End If
        Set oMuestra = Nothing
        Me.MousePointer = 0
        
    End If

   On Error GoTo 0
   Exit Sub

cmdmanual_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmanual_Click of Formulario frmPrevisualizar"

End Sub

Private Sub cmdRevisar_Click()
    If MsgBox("Va a dar por REVISADA la muestra. ¿Desea Continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    Me.MousePointer = 11
    Dim oMuestra As New clsMuestra
    oMuestra.Revisar (PK)
    Espera (2)
    Dim r As Integer
    r = 1
    Do While r = 1
        Espera (1)
        r = verificar_impresion(PK)
    Loop
    If r = 0 Then
        MsgBox "Muestra revisada correctamente.", vbOKOnly + vbInformation, App.Title
        mostrar_pdf PK, 0
    Else
        MsgBox "La generación del documento no se ha realizado correctamente.", vbExclamation, App.Title
    End If
    Set oMuestra = Nothing
    Me.MousePointer = 0
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdImprimir_Click()
    On Error GoTo fallo
    pdf1.printWithDialog
    Dim oMuestra As New clsMuestra
    oMuestra.informar_impresion PK, USUARIO.getID_EMPLEADO
    Set oMuestra = Nothing
    chkimpresa.Value = Checked
    chkimpresa.BackColor = &HC0FFFF
    Exit Sub
fallo:
    MsgBox "Error al imprimir la muestra. " & Err.Description, vbCritical, App.Title
End Sub
Private Sub cmdmail_Click()
    On Error GoTo fallo
    
    If cmdRevisar.visible = True Then
        MsgBox "No se puede enviar el informe sin estar REVISADO.", vbCritical, App.Title
        Exit Sub
    End If
    
    Me.MousePointer = 11
    enviar_informe PK, txtediciongen, Me.Hwnd
    Dim oMuestra As New clsMuestra
    oMuestra.informar_correo PK, USUARIO.getID_EMPLEADO
    Set oMuestra = Nothing
    chkCorreo.Value = Checked
    chkCorreo.BackColor = &HC0FFFF
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al enviar la muestra por correo. " & Err.Description, vbCritical, App.Title
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    On Error GoTo fallo
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (PK)
    'M1053-I
    If oMuestra.getIPA <> 0 Then
        MsgBox "ATENCION : La muestra esta marcada como IPA.", vbExclamation, App.Title
    End If
    'M1053-F
    cmdImprimir.Enabled = True
    cmdmail.Enabled = True
    cmdInforme.Enabled = True
    ' Edicion de la muestra
    txtediciongen = oMuestra.getULT_EDICION_IMP
    txtedicion = oMuestra.getULT_EDICION_IMP + 1
    If oMuestra.getIMPRESA = 0 Then
        chkimpresa.Value = Unchecked
        chkimpresa.BackColor = &HFF&
    End If
    If oMuestra.getENVIADO_CORREO = 0 Then
        chkCorreo.Value = Unchecked
        chkCorreo.BackColor = &HFF&
    End If
    ' Mostrar pdf
    If oMuestra.getCERRADA <> 1 Then
        cmdManual.Enabled = False
        pdf1.LoadFile ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\cerrada.pdf"
        pdf1.setShowToolbar False
        cmdPrev.Enabled = True
        cmdInforme.Enabled = False
        cmdImprimir.Enabled = False
        cmdmail.Enabled = False
    Else
        If oMuestra.getULT_EDICION_IMP = 0 Then
            pdf1.LoadFile ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\no.pdf"
            pdf1.setShowToolbar False
            cmdPrev.Enabled = True
            cmdImprimir.Enabled = False
            cmdmail.Enabled = False
            cmdInforme.Enabled = False
        Else
            If oMuestra.getULT_EDICION_IMP > 1 Then
                cargar_ediciones
            End If
            mostrar_pdf PK, oMuestra.getULT_EDICION_IMP
        End If
    End If
    If oMuestra.getANULADA <> 0 Then
        cmdInforme.Enabled = False
        cmdmail.Enabled = False
        cmdPrev.Enabled = False
    End If
    cargar_adjuntos
    permisos
    If oMuestra.getREVISION_USUARIO <> 0 Then
        cmdRevisar.visible = False
    End If
    ' Si es un ensayo de eficacia de tiempo, incluimos el boton de Anexo
    cmdAnexo.visible = False
    If oMuestra.getANALISIS_MODIFICADO = 2 Then
        Dim oCE_R As New clsCe_recepcion
        oCE_R.Carga PK
        Dim oCE_T As New clsCe_tipos_ensayos
        oCE_T.Carga oCE_R.getTIPO_ENSAYO_ID
        If IsNumeric(oCE_T.getHORAS) Then
            cmdAnexo.visible = True
        End If
    End If
    Exit Sub
fallo:
    Exit Sub
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Ficheros Adjuntos...", lista.Width, lvwColumnLeft
        .Add , , "Ruta", 1, lvwColumnCenter
    End With
    With listaEdiciones.ColumnHeaders
        .Add , , "Ed.", 300, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Usuario", 1200, lvwColumnLeft
        .Add , , "Motivo", 1, lvwColumnCenter
    End With
End Sub
Private Sub mostrar_pdf(MUESTRA As Long, EDICION As Integer)
    Dim destino As String
    Dim op As New clsParametros
    Dim odoc As New clsDocumentacion
    destino = odoc.CargarInforme(MUESTRA, EDICION, False, False)
    If Dir(destino) <> "" Then
        pdf1.setShowToolbar True
        op.Carga PARAM_ZOOM_INFORMES, ""
        pdf1.setZoom op.getVALOR
        pdf1.LoadFile destino
'        pdf1.setView ("Fit")
'        pdf1.execCommand "DocZoomTo(char* fullPath, char* zoomType, int scale)"
    End If
    Set odoc = Nothing
End Sub
Private Sub permisos()
    If USUARIO.getPER_EDICION = True Then
        txtedicion.Locked = False
    Else
        txtedicion.Locked = True
    End If
    If USUARIO.getPER_REVISION = True Then
        cmdRevisar.visible = True
    Else
        cmdRevisar.visible = False
    End If
End Sub
Private Sub cargar_adjuntos()
'    Dim m As Long
'    Dim rs As ADODB.RecordSet
'    lista.ListItems.Clear
'    Dim oMuestra_Adjunto As New clsMuestras_adjuntos
'    Set rs = oMuestra_Adjunto.Listado(PK)
'    If rs.RecordCount > 0 Then
'        Do
'            With lista.ListItems.Add(, , rs(0))
'                 .SubItems(1) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\" & rs(3) & "\" & rs(0)
'            End With
'            rs.MoveNext
'        Loop Until rs.EOF
'    End If
'    Set rs = Nothing
'    Set oMuestra_Adjunto = Nothing
End Sub

Private Sub Form_Resize()
    pdf1.Left = 0
    pdf1.top = 0
    pdf1.Width = Me.Width - 170
    pdf1.Height = Me.Height - (Frame2.Height + 500)
    
    Frame2.top = pdf1.Height + 100
    frmEdiciones.top = pdf1.Height + 100
    Frame4.top = pdf1.Height + 100
    Frame4.Left = pdf1.Width - Frame4.Width
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        Dim destino As String
        destino = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        On Error GoTo fallo
        If Dir(destino) <> "" Then
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbNormalFocus)
        End If
    End If
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento asociado.", vbCritical, App.Title
End Sub

Private Sub cargar_ediciones()
    frmEdiciones.visible = True
    tool

    Dim m As Long
    Dim rs As ADODB.Recordset
    listaEdiciones.ListItems.Clear
    Dim oMuestra As New clsMuestra
    Dim oMe As New clsMuestras_ediciones
    Set rs = oMe.Listado(PK)
    If rs.RecordCount > 0 Then
        Do
            With listaEdiciones.ListItems.Add(, , rs("EDICION"))
                 .SubItems(1) = Format(rs("FECHA"), "dd-mm-yyyy")
                 .SubItems(2) = rs("USUARIO_ID")
                 .SubItems(3) = rs("OBSERVACIONES")
            End With
            rs.MoveNext
        Loop Until rs.EOF
        ' Carga la edición inicial
        oMuestra.CargaMuestra PK
        Dim oUsuario As New clsUsuarios
        oUsuario.CARGAR oMuestra.getCERRADA_USUARIO
        With listaEdiciones.ListItems.Add(, , 1)
            .SubItems(1) = Format(oMuestra.getFECHA_CIERRE, "dd-mm-yyyy")
            .SubItems(2) = oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
            .SubItems(3) = "Cierre"
        End With
    End If
    Set rs = Nothing
    Set oMuestra = Nothing
End Sub
Private Sub tool()
   On Error GoTo tool_Error

   With m_cTT
    ' Creamos el toolTip pasandole el nombre del Formulario
    Call .Create(Me)
    .MaxTipWidth = 600
    .Margin(ttMarginBottom) = 7
    .Margin(ttMarginTop) = 7
    .Margin(ttMarginLeft) = 5
    .Margin(ttMarginRight) = 5
    .DelayTime(ttDelayShow) = 10000
    ' Agregamos un ToolTip
    'Nota: solo es valido usar controles que posean HWND
    .AddTool listaEdiciones
   End With

   On Error GoTo 0
   Exit Sub

tool_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tool of Formulario frmMuestraPendientesFacturacion"
End Sub

Private Sub listaEdiciones_Click()
    On Error Resume Next
    mostrar_pdf PK, CInt(listaEdiciones.ListItems(listaEdiciones.selectedItem.Index).Text)
    Dim texto As String
    texto = "Creada por : " & listaEdiciones.ListItems(listaEdiciones.selectedItem.Index).SubItems(2) & " el día : " & listaEdiciones.ListItems(listaEdiciones.selectedItem.Index).SubItems(1)
    texto = texto & vbNewLine & vbNewLine & "Motivo : " & listaEdiciones.ListItems(listaEdiciones.selectedItem.Index).SubItems(3)
    m_cTT.ToolText(listaEdiciones) = texto
End Sub

