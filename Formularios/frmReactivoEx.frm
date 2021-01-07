VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmReactivoEx 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reactivo Externo"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "frmReactivoEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFds 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver FDS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4905
      Picture         =   "frmReactivoEx.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7965
      Width           =   1140
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
      Height          =   2145
      Left            =   45
      TabIndex        =   20
      Top             =   5685
      Width           =   8220
      Begin VB.CommandButton cmdAdd2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1740
         Width           =   810
      End
      Begin VB.CommandButton cmdQ2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7230
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1740
         Width           =   885
      End
      Begin MSDataListLib.DataCombo cmbFrases 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1725
         Width           =   6165
         _ExtentX        =   10874
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
         Height          =   1410
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   2487
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
      Height          =   1980
      Left            =   45
      TabIndex        =   18
      Top             =   3645
      Width           =   8220
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1605
         Width           =   810
      End
      Begin VB.CommandButton cmdQ1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7230
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1605
         Width           =   885
      End
      Begin MSDataListLib.DataCombo cmbPictogramas 
         Height          =   315
         Left            =   135
         TabIndex        =   5
         Top             =   1575
         Width           =   6195
         _ExtentX        =   10927
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
         Height          =   1230
         Left            =   135
         TabIndex        =   4
         Top             =   285
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   2170
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   6105
      Picture         =   "frmReactivoEx.frx":0AEF
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   7215
      Picture         =   "frmReactivoEx.frx":0DF9
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7965
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3180
      Left            =   30
      TabIndex        =   14
      Top             =   375
      Width           =   8205
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Buscar código de Panreac en la web"
         Height          =   285
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2790
         Width           =   3975
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1620
         TabIndex        =   3
         Top             =   2745
         Width           =   2280
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   2
         Left            =   1620
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1725
         Width           =   6495
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1620
         TabIndex        =   0
         Top             =   225
         Width           =   6465
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   1
         Left            =   1620
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   630
         Width           =   6480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "FDS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   2790
         Width           =   390
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Seguridad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   165
         TabIndex        =   19
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   17
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Almacenaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   195
         TabIndex        =   15
         Top             =   1020
         Width           =   1125
      End
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   945
      Index           =   0
      Left            =   60
      Stretch         =   -1  'True
      Top             =   7875
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo Reactivo Externo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   30
      TabIndex        =   16
      Top             =   15
      Width           =   8190
   End
End
Attribute VB_Name = "frmReactivoEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdBuscar_Click()
    On Error GoTo fallo
'    If txtDatos(3) <> "" Then
        Dim iret As Long
        iret = ShellExecute(Me.Hwnd, vbNullString, "http://www.panreac.com/new/esp/catalogo/catalogo01.htm", vbNullString, "c:", SW_SHOWNORMAL)
'    End If
    Exit Sub
fallo:
    error_grave ("Error al abrir la página web : " & Err.Description)
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdFds_Click()
    On Error GoTo fallo
    If txtDatos(3) <> "" Then
        Dim iret As Long
        iret = ShellExecute(Me.Hwnd, vbNullString, "http://www.panreac.com/new/esp/fds/ESP/X" & Trim(txtDatos(3)) & ".htm", vbNullString, "c:", SW_SHOWNORMAL)
    End If
    Exit Sub
fallo:
    error_grave ("Error al abrir la página web : " & Err.Description)
End Sub

Private Sub cmdok_Click()
    If validar = True Then
      On Error GoTo fallo
      Dim c As String
      Dim ore As New clsTipos_reactivo_ex
      With ore
          .setNOMBRE = txtDatos(0)
          .setALMACENAJE = txtDatos(1)
          .setSEGURIDAD = txtDatos(2)
          .setFDS = txtDatos(3)
      End With
      If greactivoex = 0 Then
        If MsgBox("Va a introducir un nuevo Reactivo Externo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            greactivoex = ore.Insertar
            If greactivoex = 0 Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el Reactivo Externo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If ore.Modificar(greactivoex) = False Then
                Exit Sub
            Else
                ' Borrar Pictogramas
                c = "delete from reactivos_ex_pictogramas where reactivo_ex_id = " & greactivoex
                execute_bd c
                ' Borrar Frases
                c = "delete from reactivos_ex_frasesrys where reactivo_ex_id = " & greactivoex
                execute_bd c
            End If
        Else
            Exit Sub
        End If
      End If
      ' Insertar pictogramas
      For i = 1 To lpictogramas.ListItems.Count
        c = "insert into reactivos_ex_pictogramas values(" & greactivoex & "," & lpictogramas.ListItems(i).SubItems(1) & ")"
        execute_bd c
      Next
      ' Insertar frases
      For i = 1 To lfrases.ListItems.Count
        c = "insert into reactivos_ex_frasesrys values(" & greactivoex & "," & lfrases.ListItems(i).SubItems(2) & ")"
        execute_bd c
      Next
      If greactivoex = 0 Then
          MsgBox "El Reactivo Externo se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El Reactivo Externo se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave ("Error al insertar el reactivo externo : " & Err.Description)
End Sub

Private Sub cmdQ1_Click()
    If lpictogramas.SelectedItem.Index > 0 Then
        lpictogramas.ListItems.Remove lpictogramas.SelectedItem.Index
    End If
End Sub

Private Sub cmdQ2_Click()
    If lfrases.SelectedItem.Index > 0 Then
        lfrases.ListItems.Remove lfrases.SelectedItem.Index
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_combo cmbPictogramas, New clsPictogramas
    cargar_combo cmbFrases, New clsFrasesrys
    With lpictogramas.ColumnHeaders.Add(, , "Pictograma", 7300, lvwColumnLeft)
        .Tag = "Pictograma"
    End With
    With lpictogramas.ColumnHeaders.Add(, , "ID", 400, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lfrases.ColumnHeaders.Add(, , "Código", 1000, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lfrases.ColumnHeaders.Add(, , "Frase", 6300, lvwColumnLeft)
        .Tag = "Frase"
    End With
    With lfrases.ColumnHeaders.Add(, , "ID", 400, lvwColumnCenter)
        .Tag = "ID"
    End With
    If greactivoex <> 0 Then
        Label1(2) = "Modificación de Reactivo Externo"
        Label1(2).BackColor = &H80C0FF
        cargar_ReactivoEx
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar_ReactivoEx()
    On Error Resume Next
    Dim ore As New clsTipos_reactivo_ex
    Dim rs As ADODB.Recordset
    With ore
     .cargar (CLng(greactivoex))
     txtDatos(0) = .getNOMBRE
     txtDatos(1) = .getALMACENAJE
     txtDatos(2) = .getSEGURIDAD
     txtDatos(3) = .getFDS
     ' Pictogramas
     Dim Index As Integer
     Set rs = .Listado_Pictogramas(CLng(greactivoex))
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
           img(Index).Visible = True
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
     Set rs = .Listado_Frases(CLng(greactivoex))
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

