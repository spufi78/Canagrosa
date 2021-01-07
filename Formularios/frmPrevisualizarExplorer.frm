VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmPrevisualizarExplorer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Previsualizar Datos de Informe"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12105
   Icon            =   "frmPrevisualizarExplorer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4200
      Left            =   1395
      TabIndex        =   15
      Top             =   2070
      Width           =   3750
      ExtentX         =   6615
      ExtentY         =   7408
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
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
      Left            =   9945
      TabIndex        =   1
      Top             =   7200
      Value           =   1  'Checked
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
      Left            =   9945
      TabIndex        =   14
      Top             =   6975
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox chkCabecera 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ocultar cabecera"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9945
      TabIndex        =   13
      Top             =   5535
      Width           =   2055
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
      Left            =   10260
      Picture         =   "frmPrevisualizarExplorer.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Generar una nueva edición del informe"
      Top             =   5895
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ultima Ed. generada"
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
      Left            =   9945
      TabIndex        =   7
      Top             =   60
      Width           =   2130
      Begin VB.CommandButton cmdmail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-mail"
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
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Enviar informe de la ultima edición generada por E-mail"
         Top             =   1560
         Width           =   1545
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
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   705
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
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
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir ultima edición generada"
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label Label2 
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
         TabIndex        =   10
         Top             =   315
         Width           =   675
      End
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
      Left            =   9945
      TabIndex        =   2
      Top             =   2790
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
         Picture         =   "frmPrevisualizarExplorer.frx":0420
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmPrevisualizarExplorer.frx":171A
         Style           =   1  'Graphical
         TabIndex        =   5
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
         TabIndex        =   3
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
         TabIndex        =   4
         Top             =   315
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7695
      Width           =   1455
   End
End
Attribute VB_Name = "frmPrevisualizarExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkCabecera_Click()
    mostrar_pdf (gmuestra)
End Sub

Private Sub chkcorreo_Click()
'    MsgBox "Esta marca es solo informativa. No se puede modificar.", vbInformation, App.Title

End Sub

Private Sub chkimpresa_Click()
'    MsgBox "Esta marca es solo informativa. No se puede modificar.", vbInformation, App.Title
End Sub

Private Sub cmdacrobat_Click()
    abrir_pdf (gmuestra)
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo fallo
    Pdf1.printWithDialog
    Dim omuestra As New clsMuestra
    omuestra.informar_impresion gmuestra, USUARIO.getID_EMPLEADO
    Set omuestra = Nothing
    chkimpresa.value = Checked
    chkimpresa.BackColor = &HC0FFFF
    Exit Sub
fallo:
    MsgBox "Error al imprimir la muestra. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdInforme_Click()
    On Error GoTo fallo
    If USUARIO.getPER_EDICION = False Then
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
         Dim omuestra As New clsMuestra
         omuestra.CargaMuestra (gmuestra)
         If omuestra.getULT_EDICION_IMP <> 0 Then
            If MsgBox("La muestra tiene " & omuestra.getULT_EDICION_IMP & " edición/es impresas. ¿Generar una nueva edición?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
         End If
         ' Motivo de nueva edición
         frmMotivo.Show 1
         If Trim(motivo) = "" Then
             MsgBox "Para generar una nueva edición es necesario introducir un motivo.", vbInformation, App.Title
             Exit Sub
         End If
         If omuestra.Nueva_Edicion(gmuestra, CInt(txtedicion), Trim(motivo)) = False Then
             Exit Sub
         End If
         ' Edicion
         Me.MousePointer = 11
         omuestra.modificar_edicion_impresa gmuestra, CInt(txtedicion) - 1
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
    enviar_informe gmuestra, 0, Me.hWnd
    Dim omuestra As New clsMuestra
    omuestra.informar_correo gmuestra, USUARIO.getID_EMPLEADO
    Set omuestra = Nothing
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
         Dim omuestra As New clsMuestra
         ' Edicion
         omuestra.modificar_edicion_impresa gmuestra, CInt(txtedicion) - 1
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
    On Error GoTo fallo
    Dim omuestra As New clsMuestra
    omuestra.CargaMuestra (gmuestra)
    cmdImprimir.Enabled = True
    cmdmail.Enabled = True
    cmdInforme.Enabled = True
    ' Edicion de la muestra
    txtediciongen = omuestra.getULT_EDICION_IMP
    txtedicion = omuestra.getULT_EDICION_IMP + 1
    If omuestra.getIMPRESA = 0 Then
        chkimpresa.value = Unchecked
        chkimpresa.BackColor = &HFF&
    End If
    If omuestra.getENVIADO_CORREO = 0 Then
        chkcorreo.value = Unchecked
        chkcorreo.BackColor = &HFF&
    End If
    ' Mostrar pdf
'    MsgBox Replace(ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas"), "\", "/") & "/cerrada.pdf"
'    web.Navigate Replace(ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas"), "\", "/") & "/cerrada.pdf"
 '   WebBrowser1.Navigate "j:\prueba.pdf"
'    If omuestra.getCERRADA <> 1 Then
'        Pdf1.LoadFile ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\cerrada.pdf"
'        Pdf1.setShowToolbar False
'        cmdPrev.Enabled = True
'        cmdInforme.Enabled = False
'        cmdImprimir.Enabled = False
'        cmdmail.Enabled = False
'    Else
'        If omuestra.getULT_EDICION_IMP = 0 Then
'            Pdf1.LoadFile ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\no.pdf"
'            Pdf1.setShowToolbar False
'            cmdPrev.Enabled = True
'            cmdImprimir.Enabled = False
'            cmdmail.Enabled = False
'            cmdInforme.Enabled = False
'        Else
'            mostrar_pdf (gmuestra)
'        End If
'    End If
'    If omuestra.getANULADA <> 0 Then
'        cmdInforme.Enabled = False
'        cmdmail.Enabled = False
'        cmdPrev.Enabled = False
'    End If
    permisos
    Exit Sub
fallo:
    Exit Sub
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
        Pdf1.LoadFile destino
        Pdf1.setShowToolbar True
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
    Dim omuestra As New clsMuestra
    omuestra.aumentar_edicion_impresa (MUESTRA)
    mostrar_pdf (MUESTRA)
    omuestra.disminuir_edicion_impresa (MUESTRA)
End Sub

Public Sub permisos()
    If USUARIO.getPER_EDICION = True Then
        txtedicion.Locked = False
    Else
        txtedicion.Locked = True
    End If
End Sub

