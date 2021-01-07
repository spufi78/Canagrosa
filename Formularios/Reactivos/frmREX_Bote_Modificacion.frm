VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmREX_Bote_Modificacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificación de Bote de Reactivo"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   Icon            =   "frmREX_Bote_Modificacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Recepción"
      Enabled         =   0   'False
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
      Height          =   1095
      Left            =   45
      TabIndex        =   32
      Top             =   4815
      Width           =   7890
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   900
         TabIndex        =   33
         Top             =   630
         Width           =   3450
      End
      Begin pryCombo.miCombo cmbRecepcionado 
         Height          =   330
         Left            =   900
         TabIndex        =   36
         Top             =   270
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   35
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   34
         Top             =   675
         Width           =   450
      End
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1935
      TabIndex        =   30
      Top             =   7515
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selección del Certificado Externo"
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
      Height          =   1230
      Left            =   45
      TabIndex        =   23
      Top             =   5985
      Width           =   7890
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escaner"
         Height          =   825
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   270
         Width           =   825
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   825
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   270
         Width           =   810
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   825
         Index           =   0
         Left            =   4365
         Picture         =   "frmREX_Bote_Modificacion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   270
         Width           =   810
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   825
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   540
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   270
         Width           =   3765
      End
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar"
         Height          =   825
         Left            =   6930
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ruta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   27
         Top             =   585
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7335
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7335
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   45
      TabIndex        =   13
      Top             =   630
      Width           =   7920
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Height          =   825
         Left            =   5175
         TabIndex        =   38
         Top             =   1620
         Width           =   1950
         Begin VB.OptionButton opHenkel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   315
            TabIndex        =   40
            Top             =   270
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.OptionButton opHenkel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   1080
            TabIndex        =   39
            Top             =   270
            Width           =   795
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "HENKEL"
            Height          =   195
            Index           =   9
            Left            =   585
            TabIndex        =   41
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.CheckBox chkNoConforme 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.A.U."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   5580
         TabIndex        =   3
         Top             =   1080
         Width           =   2085
      End
      Begin VB.CheckBox chkNoCaduca 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "No caduca"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3195
         TabIndex        =   5
         Top             =   1485
         Width           =   1275
      End
      Begin VB.CheckBox chkcierre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cerrar bote"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3195
         TabIndex        =   9
         Top             =   2295
         Width           =   1275
      End
      Begin VB.CheckBox chkAbrir 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abrir bote"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3195
         TabIndex        =   7
         Top             =   1890
         Width           =   1275
      End
      Begin VB.OptionButton opAceptado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   2145
         TabIndex        =   12
         Top             =   3735
         Width           =   795
      End
      Begin VB.OptionButton opAceptado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   1455
         TabIndex        =   11
         Top             =   3735
         Width           =   555
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   825
         Index           =   3
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2880
         Width           =   7620
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   1470
         TabIndex        =   1
         Top             =   630
         Width           =   6285
      End
      Begin MSComCtl2.DTPicker fcaducidad 
         Height          =   330
         Left            =   1470
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   52035585
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fcierre 
         Height          =   330
         Left            =   1470
         TabIndex        =   8
         Top             =   2250
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   52035585
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fapertura 
         Height          =   330
         Left            =   1470
         TabIndex        =   6
         Top             =   1845
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   52035585
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker frecepcion 
         Height          =   330
         Left            =   1470
         TabIndex        =   2
         Top             =   1035
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   52035585
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmREX_Bote_Modificacion.frx":0BD4
         Height          =   315
         Left            =   1470
         TabIndex        =   0
         Top             =   225
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   180
         TabIndex        =   37
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepción"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   22
         Top             =   1110
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aceptado"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   19
         Top             =   3825
         Width           =   690
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   18
         Top             =   2655
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apertura"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   17
         Top             =   1935
         Width           =   600
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cierre"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   2340
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducidad"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   15
         Top             =   1515
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   690
         Width           =   315
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3600
      Top             =   7785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modificación de Bote de Reactivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   28
      Top             =   135
      Width           =   5805
      WordWrap        =   -1  'True
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   7380
      Picture         =   "frmREX_Bote_Modificacion.frx":0C1A
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   7965
   End
End
Attribute VB_Name = "frmREX_Bote_Modificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub chkNoCaduca_Click()
    If chkNoCaduca.Value = Checked Then
        fCaducidad.Enabled = False
    Else
        fCaducidad.Enabled = True
    End If
End Sub

Private Sub cmdAdjuntar_Click()
   On Error GoTo cmdVincular_Click_Error

    If datos(0) = "" Then
        MsgBox "Por favor, indique el certificado a vincular.", vbExclamation, App.Title
        Exit Sub
    End If
    If Dir(datos(0)) = "" Then
        MsgBox "El documento vinculado no existe en la ruta.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oAdjunto As New clsAdjuntos
    Dim adjunto As Long
    With oAdjunto
'        .Eliminar TOBJETO.TOBJETO_REX_CERTIFICADOS, PK, 0, 0
        
        .setTIPO = TOBJETO.TOBJETO_REX_CERTIFICADOS
        .setCODIGO = PK
        .setCODIGO_DECODIFICADORA = 0
        .setTIPO_DOCUMENTO_ID = ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_CERTIFICADO
        .setOBSERVACIONES = ""
        .setFICHERO_NOMBRE = datos(1)
        .setFICHERO_RUTA = datos(0)
        adjunto = .Insertar(0, False)
    End With
    If adjunto > 0 Then
        ' Actualizar el codigo del adjunto
        Dim oBote As New clsBotes_ex
        oBote.setCERTIFICADO_EXTERNO = adjunto & " - " & datos(1)
        oBote.InformarRutaCertificado PK
        Set oBote = Nothing
    End If
    Set oAdjunto = Nothing

    MsgBox "Certificado vinculado correctamente.", vbInformation, App.Title
   On Error GoTo 0
   Exit Sub

cmdVincular_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVincular_Click of Formulario frmREX_Bote_Modificacion"
    

End Sub

Private Sub cmdEscaner_Click()
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = ""
        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            datos(0).Text = documento_escaner
            datos(1).Text = nombreNuevo & ".pdf"
            cmdAdjuntar_Click
        End If
    End If
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
'    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
'        datos(4).Text = cd.FileTitle 'cd.FileName  '
        datos(0).Text = cd.FileName
        datos(1).Text = cd.FileTitle
    End If
End Sub

Private Sub cmdMostrar_Click()
    On Error GoTo fallo
' M0601-I
'    If datos(0) <> "" Then
'        If Dir(datos(0)) <> "" Then
'            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & datos(0), vbMaximizedFocus)
'        End If
'    End If
    Dim oAdjunto As New clsAdjuntos
    If oAdjunto.CargarDocumentoUltimo(TOBJETO.TOBJETO_REX_CERTIFICADOS, PK, 0, True, ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_CERTIFICADO) = "" Then
        MsgBox "El certificado no esta informado.", vbInformation, App.Title
    End If
    Set oAdjunto = Nothing
' M0601-F
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title

End Sub
Private Sub datos_GotFocus(Index As Integer)
    datos(Index).BackColor = &HC0FFFF
End Sub

Private Sub datos_LostFocus(Index As Integer)
    datos(Index).BackColor = vbWhite
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        Dim obe As New clsBotes_ex
        With obe
           .setCENTRO_ID = cmbCentro.BoundText
           .setLOTE = txtDatos(2)
           .setFECHA_RECEPCION = Format(frecepcion.Value, "yyyy-mm-dd")
           If chkNoCaduca = Checked Then
             .setFECHA_CADUCIDAD = "0000-00-00"
             .setNO_CADUCA = 1
           Else
             .setFECHA_CADUCIDAD = Format(fCaducidad.Value, "yyyy-mm-dd")
             .setNO_CADUCA = 0
           End If
           .setNO_CONFORME = chkNoConforme.Value
           If chkAbrir.Value = Checked Then
                .setFECHA_APERTURA = Format(fapertura.Value, "yyyy-mm-dd")
                .setABIERTO = 1
           Else
                .setABIERTO = 0
                .setFECHA_APERTURA = "0000-00-00"
           End If
           If chkcierre.Value = Checked Then
                .setFECHA_FIN = Format(fcierre.Value, "yyyy-mm-dd")
                .setFINALIZADO = 1
           Else
                .setFECHA_FIN = "0000-00-00"
                .setFINALIZADO = 0
           End If
           .setOBSERVACIONES = txtDatos(3)
           If opAceptado(0).Value = True Then
            .setANULADO = 0
           Else
            .setANULADO = 1
           End If
           If opHenkel(0).Value = True Then
            .setHENKEL = 0
           Else
            .setHENKEL = 1
           End If
'           .setCERTIFICADO_EXTERNO = datos(0)
           If .Modificar(PK) = True Then
              MsgBox "Modificaciones realizadas correctamente.", vbOKOnly + vbInformation, App.Title
              Unload Me
           End If
        End With
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_Bote_Modificacion")
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    permisos
    llenar_combo cmbRecepcionado, New clsUsuarios, 0, frmUsuarios, ""
    cargar_combo cmbCentro, New clsCentros
    Call cargar_bote
End Sub
Private Sub cargar_bote()
    Dim oBote As New clsBotes_ex
   On Error GoTo cargar_bote_Error

    With oBote
'        If .CARGAR(gbotereactivoex) = True Then
        If .cargar(PK) = True Then
            cmbCentro.BoundText = .getCENTRO_ID
            txtDatos(2) = .getLOTE
            frecepcion = .getFECHA_RECEPCION
            If .getFECHA_CADUCIDAD = "" Or .getFECHA_CADUCIDAD = "0000-00-00" Then
                fCaducidad = Date
                chkNoCaduca.Value = Checked
            Else
                fCaducidad = .getFECHA_CADUCIDAD
                chkNoCaduca.Value = Unchecked
            End If
            chkNoConforme = .getNO_CONFORME
            If .getFECHA_APERTURA = "" Then
                fapertura.Value = Date
                chkAbrir.Value = Unchecked
            Else
                fapertura = .getFECHA_APERTURA
                chkAbrir.Value = Checked
            End If
            If .getFECHA_FIN = "" Then
                fcierre.Value = Date
                chkcierre.Value = Unchecked
            Else
                fcierre = .getFECHA_FIN
                chkcierre.Value = Checked
            End If
            txtDatos(3) = .getOBSERVACIONES
            If .getANULADO = 0 Then
                opAceptado(0).Value = True
            Else
                opAceptado(1).Value = True
            End If
            opHenkel(.getHENKEL).Value = True
            datos(0) = .getCERTIFICADO_EXTERNO
            ' Datos de recepcion
            If .getPEDIDO_BOTE_EX_ID <> 0 Then
                Dim oPed As New clsPedidos_bote_ex
                oPed.cargar .getPEDIDO_BOTE_EX_ID
                Dim rsAux As ADODB.Recordset
                Set rsAux = oPed.CARGAR_POR_PEDIDO_PROVEEDOR(oPed.getCODIGO_PEDIDO_PROVEEDOR, Year(oPed.getFECHA_PEDIDO))
                If rsAux.RecordCount > 0 Then
                    If rsAux("RECEPCION_USUARIO") <> 0 Then
                        cmbRecepcionado.MostrarElemento rsAux("RECEPCION_USUARIO")
                    End If
                    If Not IsNull(rsAux("RECEPCION_FECHA")) Then
                        txtDatos(0) = rsAux("RECEPCION_FECHA")
                    End If
                End If
            End If
        End If
    End With
    Set oBote = Nothing

   On Error GoTo 0
   Exit Sub

cargar_bote_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_bote of Formulario frmREX_Bote_Modificacion"
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbCentro.Text = "" Then
        MsgBox "Debe introducir el CENTRO.", vbInformation, App.Title
        cmbCentro.SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(2)) = "" Then
        MsgBox "Debe introducir el lote.", vbInformation, App.Title
        txtDatos(2).SetFocus
        validar = False
        Exit Function
    End If
End Function


Public Sub permisos()
    If USUARIO.getPER_PEDIDOS_REACTIVOS = False Then
        Frame2.Enabled = False
    End If
End Sub
