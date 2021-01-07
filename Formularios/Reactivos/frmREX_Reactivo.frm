VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmREX_Reactivo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Sustancias / Materiales"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   Icon            =   "frmREX_Reactivo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ficha de Seguridad PANREAC"
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
      Height          =   1005
      Left            =   45
      TabIndex        =   29
      Top             =   6750
      Width           =   7215
      Begin VB.CommandButton cmdFds 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver FDS"
         Height          =   735
         Left            =   3240
         Picture         =   "frmREX_Reactivo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton cmdBuscarPanreac 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar código de Panreac en el Catalogo"
         Height          =   735
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   180
         Width           =   2805
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   3
         Left            =   1170
         TabIndex        =   30
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de FDS"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   450
         Width           =   765
      End
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4095
      TabIndex        =   28
      Top             =   8820
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.Frame frmFDS 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ficha de Seguridad"
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
      Height          =   1005
      Left            =   45
      TabIndex        =   21
      Top             =   7785
      Width           =   7215
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar"
         Height          =   735
         Left            =   6210
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   180
         Width           =   810
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   690
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   630
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   225
         Width           =   2955
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   735
         Index           =   0
         Left            =   3645
         Picture         =   "frmREX_Reactivo.frx":0AEF
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   180
         Width           =   810
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   735
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Width           =   810
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escaner"
         Height          =   735
         Left            =   5355
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Width           =   825
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
         Top             =   450
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frases R y S"
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
      Height          =   2910
      Left            =   5895
      TabIndex        =   19
      Top             =   3780
      Width           =   5790
      Begin VB.CommandButton cmdAdd2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   285
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2505
         Width           =   810
      End
      Begin VB.CommandButton cmdQ2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar"
         Height          =   285
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2505
         Width           =   885
      End
      Begin MSDataListLib.DataCombo cmbFrases 
         Height          =   315
         Left            =   75
         TabIndex        =   9
         Top             =   2520
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSComctlLib.ListView lfrases 
         Height          =   2265
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   3995
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pictogramas"
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
      Height          =   2925
      Left            =   45
      TabIndex        =   17
      Top             =   3780
      Width           =   5790
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   285
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2505
         Width           =   810
      End
      Begin VB.CommandButton cmdQ1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar"
         Height          =   285
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2505
         Width           =   885
      End
      Begin MSDataListLib.DataCombo cmbPictogramas 
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Top             =   2475
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSComctlLib.ListView lpictogramas 
         Height          =   2175
         Left            =   90
         TabIndex        =   4
         Top             =   285
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   3836
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
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9000
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   10590
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9000
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   30
      TabIndex        =   14
      Top             =   510
      Width           =   11670
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   4
         Left            =   1170
         TabIndex        =   1
         Top             =   630
         Width           =   10335
      End
      Begin VB.CheckBox chkProbeta 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Es de tipo PROBETA"
         Height          =   195
         Left            =   135
         TabIndex        =   34
         Top             =   2925
         Width           =   3030
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   2
         Left            =   1170
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1905
         Width           =   10335
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   0
         Left            =   1170
         TabIndex        =   0
         Top             =   225
         Width           =   10335
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   825
         Index           =   1
         Left            =   1170
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1035
         Width           =   10335
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "En Certificado"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   35
         Top             =   675
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Seguridad"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   18
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   16
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Almacenaje"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   15
         Top             =   1425
         Width           =   825
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5490
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Definición de Sustancias / Materiales"
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
      Index           =   0
      Left            =   90
      TabIndex        =   20
      Top             =   90
      Width           =   3870
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   945
      Index           =   0
      Left            =   60
      Stretch         =   -1  'True
      Top             =   8820
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   13230
   End
End
Attribute VB_Name = "frmREX_Reactivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal Hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdAdd_Click()
    If cmbPictogramas.BoundText <> "" Then
        With lpictogramas.ListItems.Add(, , cmbPictogramas.Text)
            .SubItems(1) = cmbPictogramas.BoundText
        End With
    End If
End Sub

Private Sub cmdAdd2_Click()
    If cmbFrases.BoundText <> "" Then
        Dim oFrasesrys As New clsFrasesrys
        If oFrasesrys.Carga(cmbFrases.BoundText) = True Then
            With lfrases.ListItems.Add(, , oFrasesrys.getCODIGO)
                .SubItems(1) = oFrasesrys.getFRASE
                .SubItems(2) = oFrasesrys.getID_FRASE
            End With
        End If
    End If
End Sub

Private Sub cmdAdjuntar_Click()
    'M0953-I
   On Error GoTo cmdAdjuntar_Click_Error

    If datos(0) = "" Then
        MsgBox "Por favor, indique la FDS a vincular.", vbExclamation, App.Title
        Exit Sub
    End If
    If Dir(datos(0)) = "" Then
        MsgBox "El documento vinculado no existe en la ruta.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oAdjunto As New clsAdjuntos
    Dim adjunto As Long
    With oAdjunto
        .Eliminar TOBJETO.TOBJETO_REX_REACTIVO, PK, 0, 0
        
        .setTIPO = TOBJETO.TOBJETO_REX_REACTIVO
        .setCODIGO = PK
        .setCODIGO_DECODIFICADORA = 0
        .setTIPO_DOCUMENTO_ID = 9
        .setOBSERVACIONES = datos(1)
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
    'M0953-F

   On Error GoTo 0
   Exit Sub

cmdAdjuntar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntar_Click of Formulario frmREX_Reactivo"
End Sub

Private Sub cmdBuscarPanreac_Click()
    On Error GoTo fallo
'    If txtDatos(3) <> "" Then
        Dim iret As Long
        Dim catalogo As String
        catalogo = ReadINI(App.Path & "\config.ini", "Otros", "panreac_catalogo")
        iret = ShellExecute(Me.Hwnd, vbNullString, catalogo, vbNullString, "c:", SW_SHOWNORMAL)
'    End If
    Exit Sub
fallo:
    error_grave ("Error al abrir la página web : " & Err.Description)

End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEscaner_Click()
    'M0953-I
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
    'M0953-F
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    'M0953-I
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileName
        datos(1).Text = cd.FileTitle
    End If
    'M0953-F
End Sub

Private Sub cmdFds_Click()
    On Error GoTo fallo
    If txtDatos(3) <> "" Then
        Dim iret As Long
        Dim web_url As String
        
        web_url = ReadINI(App.Path & "\config.ini", "Otros", "panreac_fds")
        web_url = Replace(web_url, "$codigo$", Trim(txtDatos(3)))
        
        iret = ShellExecute(Me.Hwnd, vbNullString, web_url, vbNullString, "c:", SW_SHOWNORMAL)
    Else
        MsgBox "Debe introducir el código de FDS se Panreac.", vbExclamation, App.Title
    End If
    Exit Sub
fallo:
    error_grave ("Error al abrir la página web : " & Err.Description)
End Sub

Private Sub cmdMostrar_Click()
    'M0953-I
    Dim oAdjunto As New clsAdjuntos
   On Error GoTo CMDMOSTRAR_Click_Error

    If oAdjunto.CargarDocumentoUltimo(TOBJETO.TOBJETO_REX_REACTIVO, PK, 0, True, ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_CERTIFICADO) = "" Then
        MsgBox "La FDS no esta adjunta.", vbInformation, App.Title
    End If
    Set oAdjunto = Nothing
    'M0953-F

   On Error GoTo 0
   Exit Sub

CMDMOSTRAR_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMostrar_Click of Formulario frmREX_Reactivo"

End Sub

Private Sub cmdok_Click()
    If validar = True Then
      On Error GoTo fallo
      Dim c As String
      Dim ore As New clsTipos_reactivo_ex
      With ore
          .setNOMBRE = txtDatos(0)
          .setCERTIFICADO = txtDatos(4)
          .setALMACENAJE = txtDatos(1)
          .setSEGURIDAD = txtDatos(2)
          .setFDS = txtDatos(3)
          .setPROBETA = chkProbeta.Value
      End With
'      If greactivoex = 0 Then
      If PK = 0 Then
        If MsgBox("Va a introducir una nueva Sustancia/Material. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'            greactivoex = ore.Insertar
           PK = ore.Insertar
'            If greactivoex = 0 Then
            If PK = 0 Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar la Sustancia/Material. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'            If ore.Modificar(greactivoex) = False Then
            If ore.Modificar(PK) = False Then
                Exit Sub
            Else
                ' Borrar Pictogramas
'                c = "delete from reactivos_ex_pictogramas where reactivo_ex_id = " & greactivoex
                c = "delete from reactivos_ex_pictogramas where reactivo_ex_id = " & PK
                execute_bd c
                ' Borrar Frases
'                c = "delete from reactivos_ex_frasesrys where reactivo_ex_id = " & greactivoex
                c = "delete from reactivos_ex_frasesrys where reactivo_ex_id = " & PK
                execute_bd c
            End If
        Else
            Exit Sub
        End If
      End If
      ' Insertar pictogramas
      For i = 1 To lpictogramas.ListItems.Count
'        c = "insert into reactivos_ex_pictogramas values(" & greactivoex & "," & lpictogramas.ListItems(i).SubItems(1) & ")"
        c = "insert into reactivos_ex_pictogramas values(" & PK & "," & lpictogramas.ListItems(i).SubItems(1) & ")"
        execute_bd c
      Next
      ' Insertar frases
      For i = 1 To lfrases.ListItems.Count
'        c = "insert into reactivos_ex_frasesrys values(" & greactivoex & "," & lfrases.ListItems(i).SubItems(2) & ")"
        c = "insert into reactivos_ex_frasesrys values(" & PK & "," & lfrases.ListItems(i).SubItems(2) & ")"
        execute_bd c
      Next
      If greactivoex = 0 Then
          MsgBox "La Sustancia/Material se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "La Sustancia/Material se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave ("Error al insertar el reactivo externo : " & Err.Description)
End Sub

Private Sub cmdQ1_Click()
    If lpictogramas.selectedItem.Index > 0 Then
        lpictogramas.ListItems.Remove lpictogramas.selectedItem.Index
    End If
End Sub

Private Sub cmdQ2_Click()
    If lfrases.selectedItem.Index > 0 Then
        lfrases.ListItems.Remove lfrases.selectedItem.Index
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo cmbPictogramas, New clsPictogramas
    cargar_combo cmbFrases, New clsFrasesrys
    With lpictogramas.ColumnHeaders.Add(, , "Pictograma", 5500, lvwColumnLeft)
        .Tag = "Pictograma"
    End With
    With lpictogramas.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lfrases.ColumnHeaders.Add(, , "Código", 1000, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lfrases.ColumnHeaders.Add(, , "Frase", 4500, lvwColumnLeft)
        .Tag = "Frase"
    End With
    With lfrases.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
'    If greactivoex <> 0 Then
    If PK <> 0 Then
'        Label1(2) = "Modificación de Reactivo Externo"
'        Label1(2).BackColor = &H80C0FF
        cargar_ReactivoEx
        'M0953-I
        frmFDS.Enabled = True
        'M0953-F
    End If
End Sub

Private Sub txtDatos_Change(Index As Integer)
    If Index = 0 Then
        txtDatos(4) = txtDatos(Index)
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_ReactivoEx()
    On Error Resume Next
    Dim ore As New clsTipos_reactivo_ex
    Dim rs As ADODB.Recordset
    With ore
'     .CARGAR (CLng(greactivoex))
     .cargar (PK)
     txtDatos(0) = .getNOMBRE
     txtDatos(4) = .getCERTIFICADO
     txtDatos(1) = .getALMACENAJE
     txtDatos(2) = .getSEGURIDAD
     txtDatos(3) = .getFDS
     chkProbeta.Value = .getPROBETA
     'M0953-I
     Dim oAdjunto As New clsAdjuntos
     If oAdjunto.cargarDatos(TOBJETO.TOBJETO_REX_REACTIVO, PK, 0, 0) = True Then
        datos(0) = oAdjunto.getOBSERVACIONES
     End If
     'M0953-F
     ' Pictogramas
     Dim Index As Integer
'     Set rs = .Listado_Pictogramas(CLng(greactivoex))
     Set rs = .Listado_Pictogramas(PK)
     If rs.RecordCount <> 0 Then
        Do
           With lpictogramas.ListItems.Add(, , rs(0))
            .SubItems(1) = rs(1)
           End With
           ' Imagen
           If Index > 0 Then
            Load img(Index)
            img(Index).Left = img(Index - 1).Left + img(Index).Width + 10
           End If
           img(Index).visible = True
           img(Index).Picture = Nothing
           If rs(2) <> "" Then
                Dim ruta As String
                ruta = ReadINI(App.Path + "\config.ini", "Otros", "Pictogramas") & "\" & rs(2)
                If Dir(ruta) <> "" Then
                    Set img(Index).Picture = LoadPicture(ruta)
                End If
           End If
           Index = Index + 1
           rs.MoveNext
        Loop Until rs.EOF
     End If
     ' Frases R y S
'     Set rs = .Listado_Frases(CLng(greactivoex))
     Set rs = .Listado_Frases(PK)
     If rs.RecordCount <> 0 Then
        Do
           With lfrases.ListItems.Add(, , rs(0))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
           End With
           rs.MoveNext
        Loop Until rs.EOF
     End If
    End With
    Set ore = Nothing
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre al Reactivo.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function

