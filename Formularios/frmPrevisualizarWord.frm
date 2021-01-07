VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrevisualizarWord 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Previsualizar Datos de Informe"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8310
   Icon            =   "frmPrevisualizarWord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   135
      TabIndex        =   17
      Top             =   3960
      Width           =   4920
      Begin MSComctlLib.ListView lista 
         Height          =   1065
         Left            =   90
         TabIndex        =   18
         Top             =   225
         Width           =   4755
         _ExtentX        =   8387
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
      Left            =   90
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   3165
      Begin MSComctlLib.ListView listaEdiciones 
         Height          =   1065
         Left            =   90
         TabIndex        =   16
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
      Left            =   6705
      Picture         =   "frmPrevisualizarWord.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Generar una nueva edición del informe"
      Top             =   270
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
      Left            =   6030
      TabIndex        =   1
      Top             =   1575
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
         Picture         =   "frmPrevisualizarWord.frx":0420
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmPrevisualizarWord.frx":171A
         Style           =   1  'Graphical
         TabIndex        =   4
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
         TabIndex        =   2
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
         TabIndex        =   3
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
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   5415
      Begin VB.CheckBox chkCabecera 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ocultar cabecera"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   270
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
         Left            =   3240
         TabIndex        =   13
         Top             =   540
         Value           =   1  'Checked
         Width           =   2055
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
         Left            =   3240
         TabIndex        =   12
         Top             =   765
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdmail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-mail"
         Height          =   870
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Enviar informe de la ultima edición generada por E-mail"
         Top             =   270
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
         Height          =   360
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   705
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   870
         Left            =   945
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir ultima edición generada"
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   405
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   1095
      Left            =   7290
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   960
   End
End
Attribute VB_Name = "frmPrevisualizarWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_cTT As New cTooltip
Private Sub chkCabecera_Click()
    mostrar_pdf (gmuestra)
End Sub
Private Sub cmdacrobat_Click()
    abrir_pdf (gmuestra)
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdImprimir_Click()
    On Error GoTo fallo
'    Pdf1.printWithDialog
    Dim oMuestra As New clsMuestra
    oMuestra.informar_impresion gmuestra, usuario.getID_EMPLEADO
    Set oMuestra = Nothing
    chkimpresa.value = Checked
    chkimpresa.BackColor = &HC0FFFF
    Exit Sub
fallo:
    MsgBox "Error al imprimir la muestra. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdInforme_Click()
    On Error GoTo fallo
    If usuario.getPER_EDICION = False Then
        MsgBox "Su usuario no tiene permisos para generar nuevas ediciones. Contacte con su gerente.", vbExclamation, App.Title
        Exit Sub
    End If
    ' Verificar si la edición anterior se genero
    Dim destino As String
    destino = NOMBRE_DOCUMENTO(gmuestra, True) & ".pdf"
    If Dir(destino) = "" Then
        MsgBox "La edición anterior falló al generarse, por lo que no pueden generarse nuevas ediciones. Contacte con mantenimiento.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("Va a generar el informe de la muestra. ¿Desea Continuar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
         Dim oMuestra As New clsMuestra
         oMuestra.CargaMuestra (gmuestra)
         If oMuestra.getULT_EDICION_IMP <> 0 Then
            If MsgBox("La muestra tiene " & oMuestra.getULT_EDICION_IMP & " edición/es impresas. ¿Generar una nueva edición?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
         End If
         ' Motivo de nueva edición
         frmMotivo.Show 1
         If Trim(motivo) = "" Then
             MsgBox "Para generar una nueva edición es necesario introducir un motivo.", vbInformation, App.Title
             Exit Sub
         End If
         If oMuestra.Nueva_Edicion(gmuestra, CInt(txtedicion), Trim(motivo)) = False Then
             Exit Sub
         End If
         ' Edicion
         Me.MousePointer = 11
         oMuestra.modificar_edicion_impresa gmuestra, CInt(txtedicion) - 1
         If imprimir(gmuestra, 1, True) = True Then
            Form_Load
         Else
            Me.MousePointer = 0
            MsgBox "Se ha producido un error al generar el documento.", vbCritical, App.Title
         End If
    End If
    Me.MousePointer = 0
    Me.SetFocus
    Exit Sub
fallo:
    MsgBox "Se ha producido un error al generar el documento.", vbCritical, App.Title
End Sub

Private Sub cmdmail_Click()
    On Error GoTo fallo
    Me.MousePointer = 11
    enviar_informe gmuestra, 0, Me.Hwnd
    Dim oMuestra As New clsMuestra
    oMuestra.informar_correo gmuestra, usuario.getID_EMPLEADO
    Set oMuestra = Nothing
    chkcorreo.value = Checked
    chkcorreo.BackColor = &HC0FFFF
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al enviar la muestra por correo. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdPrev_Click()
    If MsgBox("Va a generar la previsualización de la muestra. ¿Desea Continuar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
         Me.MousePointer = 11
         Dim oMuestra As New clsMuestra
         ' Edicion
         oMuestra.modificar_edicion_impresa gmuestra, CInt(txtedicion) - 1
         If imprimir(gmuestra, 3, True) = True Then
            Me.MousePointer = 0
            prev_pdf (gmuestra)
            cmdImprimir.Enabled = True
         Else
            Me.MousePointer = 0
            MsgBox "Se ha producido un error al generar el documento.", vbCritical, App.Title
         End If
    End If
    Me.MousePointer = 0
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    On Error GoTo fallo
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (gmuestra)
    cmdImprimir.Enabled = True
    cmdmail.Enabled = True
    cmdInforme.Enabled = True
    ' Edicion de la muestra
    txtediciongen = oMuestra.getULT_EDICION_IMP
    txtedicion = oMuestra.getULT_EDICION_IMP + 1
    If oMuestra.getIMPRESA = 0 Then
        chkimpresa.value = Unchecked
        chkimpresa.BackColor = &HFF&
    End If
    If oMuestra.getENVIADO_CORREO = 0 Then
        chkcorreo.value = Unchecked
        chkcorreo.BackColor = &HFF&
    End If
    ' Mostrar pdf
    If oMuestra.getCERRADA <> 1 Then
'        Pdf1.LoadFile ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\cerrada.pdf"
'        Pdf1.setShowToolbar False
        cmdPrev.Enabled = True
        cmdInforme.Enabled = False
        cmdImprimir.Enabled = False
        cmdmail.Enabled = False
    Else
        If oMuestra.getULT_EDICION_IMP = 0 Then
'            Pdf1.LoadFile ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\no.pdf"
'            Pdf1.setShowToolbar False
            cmdPrev.Enabled = True
            cmdImprimir.Enabled = False
            cmdmail.Enabled = False
            cmdInforme.Enabled = False
        Else
            mostrar_pdf (gmuestra)
            If oMuestra.getULT_EDICION_IMP > 1 Then
                cargar_ediciones
            End If
        End If
    End If
    If oMuestra.getANULADA <> 0 Then
        cmdInforme.Enabled = False
        cmdmail.Enabled = False
        cmdPrev.Enabled = False
    End If
    cargar_adjuntos
    permisos
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
        .Add , , "Edición", 1000, lvwColumnLeft
        .Add , , "Fecha", 800, lvwColumnCenter
        .Add , , "Archivo", 1, lvwColumnCenter
        .Add , , "Motivo", 1, lvwColumnCenter
        .Add , , "Usuario", 1, lvwColumnCenter
    End With
End Sub
Private Sub mostrar_pdf(MUESTRA As Long)
    Dim destino As String
    If chkCabecera.value = Checked Then
        destino = NOMBRE_DOCUMENTO(MUESTRA, True) & "--.pdf"
        If Dir(destino) = "" Then
            destino = NOMBRE_DOCUMENTO(MUESTRA, True) & ".pdf"
        End If
    Else
        destino = NOMBRE_DOCUMENTO(MUESTRA, True) & ".pdf"
    End If
    If Dir(destino) <> "" Then
'        Pdf1.LoadFile destino
'        Pdf1.setShowToolbar True
    End If
End Sub
Private Sub abrir_pdf(MUESTRA As Long)
    Dim destino As String
    On Error GoTo fallo
    If chkCabecera.value = Checked Then
        destino = NOMBRE_DOCUMENTO(MUESTRA, True) & "--.pdf"
        If Dir(destino) = "" Then
            destino = NOMBRE_DOCUMENTO(MUESTRA, True) & ".pdf"
        End If
    Else
        destino = NOMBRE_DOCUMENTO(MUESTRA, True) & ".pdf"
    End If
    If Dir(destino) <> "" Then
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
    End If
    cmdcancel_Click
    Exit Sub
fallo:
    MsgBox "Error al abrir el acrobat reader.", vbCritical, App.Title
End Sub
Private Sub prev_pdf(MUESTRA As Long)
    Dim oMuestra As New clsMuestra
    oMuestra.aumentar_edicion_impresa (MUESTRA)
    mostrar_pdf (MUESTRA)
    oMuestra.disminuir_edicion_impresa (MUESTRA)
End Sub

Public Sub permisos()
    If usuario.getPER_EDICION = True Then
        txtedicion.Locked = False
    Else
        txtedicion.Locked = True
    End If
End Sub
Private Sub cargar_adjuntos()
'    Dim m As Long
'    Dim rs As ADODB.RecordSet
'    lista.ListItems.Clear
'    Dim oMuestra_Adjunto As New clsMuestras_adjuntos
'    Set rs = oMuestra_Adjunto.Listado(gmuestra)
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
    frmEdiciones.Visible = True
    tool

    Dim m As Long
    Dim rs As ADODB.RecordSet
    listaEdiciones.ListItems.Clear
    Dim oMuestra As New clsMuestra
    Set rs = oMuestra.Listado_ediciones(gmuestra)
    If rs.RecordCount > 0 Then
        Do
            With listaEdiciones.ListItems.Add(, , "Edición " & rs("EDICION"))
                 .SubItems(1) = Format(rs("FECHA"), "dd-mm-yyyy")
                 .SubItems(2) = ""
                 .SubItems(3) = rs("OBSERVACIONES")
                 .SubItems(4) = rs("USUARIO")
            End With
            rs.MoveNext
        Loop Until rs.EOF
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
    Dim texto As String
    texto = "Creada por : " & listaEdiciones.ListItems(listaEdiciones.selectedItem.Index).SubItems(4) & " el día : " & listaEdiciones.ListItems(listaEdiciones.selectedItem.Index).SubItems(1)
    texto = texto & vbNewLine & vbNewLine & "Motivo : " & listaEdiciones.ListItems(listaEdiciones.selectedItem.Index).SubItems(3)
    m_cTT.ToolText(listaEdiciones) = texto
End Sub
