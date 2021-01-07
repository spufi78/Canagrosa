VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmAlodine_Listado_Lotes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Lotes de Alodine"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13260
   Icon            =   "frmAlodine_Listado_Lotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   13260
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas (EN)"
      Height          =   870
      Index           =   1
      Left            =   6975
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8370
      Width           =   1185
   End
   Begin VB.CommandButton cmdClientes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clientes"
      Height          =   870
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8370
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   45
      TabIndex        =   14
      Top             =   675
      Width           =   13155
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   8160
         TabIndex        =   2
         Top             =   540
         Width           =   3585
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   825
         Left            =   12105
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   1005
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   3420
         TabIndex        =   1
         Top             =   180
         Width           =   1770
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   855
         TabIndex        =   0
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10620
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   840
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   300
         Left            =   11460
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   2006
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196613
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2006
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSDataListLib.DataCombo cmbProducto 
         Height          =   330
         Left            =   4320
         TabIndex        =   18
         Top             =   90
         Visible         =   0   'False
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   855
         TabIndex        =   23
         Top             =   540
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fecha_i 
         Height          =   330
         Left            =   8145
         TabIndex        =   25
         Top             =   180
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   51773441
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_f 
         Height          =   330
         Left            =   10350
         TabIndex        =   26
         Top             =   180
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   51773441
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta el"
         Height          =   240
         Index           =   6
         Left            =   9630
         TabIndex        =   28
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde el"
         Height          =   240
         Index           =   4
         Left            =   7380
         TabIndex        =   27
         Top             =   225
         Width           =   675
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   24
         Top             =   585
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   21
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   0
         Left            =   7380
         TabIndex        =   20
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   195
         Left            =   10170
         TabIndex        =   17
         Top             =   225
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.CheckBox chkprev 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Previsualizar"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9900
      TabIndex        =   13
      Top             =   8730
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas (ES)"
      Height          =   870
      Index           =   0
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8370
      Width           =   1185
   End
   Begin VB.CommandButton cmdAlbaran 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Albaranes"
      Height          =   870
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8370
      Width           =   1185
   End
   Begin VB.CommandButton cmdCertificado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificados"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8370
      Width           =   1185
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11925
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8370
      Width           =   1290
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8370
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8370
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8370
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6615
      Left            =   60
      TabIndex        =   5
      Top             =   1710
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   11668
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Lotes Planificados de Alodine"
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
      TabIndex        =   16
      Top             =   90
      Width           =   4275
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12645
      Picture         =   "frmAlodine_Listado_Lotes.frx":1272
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Especifique los datos necesarios para localizar un LOTE"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   360
      Width           =   3975
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   45
      Top             =   0
      Width           =   13185
   End
End
Attribute VB_Name = "frmAlodine_Listado_Lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbClientes_change()
    cargar_lista
End Sub

Private Sub cmdAlbaran_Click()
   On Error GoTo cmdAlbaran_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim oDoc As New clsDocumentacion
        If oDoc.CargarAlodine(lista.ListItems(lista.selectedItem.Index).SubItems(7), 2, True) = "" Then
            MsgBox "No se localiza el documento de alodine para imprimir.", vbCritical, App.Title
        End If
        Set oDoc = Nothing
    
'        Dim documento As String
'        documento = nombre_alodine(lista.ListItems(lista.selectedItem.Index).SubItems(7)) & ".pdf"
'        If Dir(documento) = "" Then
'            MsgBox "No se localiza el documento de alodine para imprimir.", vbCritical, App.Title
'        Else
'            Shell "rundll32.exe url.dll,FileProtocolHandler " & documento, vbMaximizedFocus
'        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdAlbaran_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAlbaran_Click of Formulario frmAlodine_Listado_Lotes"

End Sub

Private Sub cmdCertificado_Click()
   On Error GoTo cmdCertificado_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim oDoc As New clsDocumentacion
        If oDoc.CargarAlodine(lista.ListItems(lista.selectedItem.Index).SubItems(7), 1, True) = "" Then
            MsgBox "No se localiza el documento de alodine para imprimir.", vbCritical, App.Title
        End If
        Set oDoc = Nothing
'        Dim documento As String
'        documento = nombre_alodine(lista.ListItems(lista.selectedItem.Index).SubItems(7)) & tipo & " CERT.pdf"
'        If Dir(documento) = "" Then
'            MsgBox "No se localiza el documento de alodine para imprimir.", vbCritical, App.Title
'        Else
'            Shell "rundll32.exe url.dll,FileProtocolHandler " & documento, vbMaximizedFocus
'        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdCertificado_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCertificado_Click of Formulario frmAlodine_Listado_Lotes"

End Sub

Private Sub cmdClientes_Click()
    If lista.ListItems.Count > 0 Then
        glote = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        Dim oAL As New clsAlodine_lotes
        oAL.Carga glote
        gAlodine = oAL.getALODINE_ID
        frmAlodine_Clientes.Show 1
        glote = 0
        gAlodine = 0
    End If
End Sub


Private Sub cmdetiqueta_Click(Index As Integer)
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim objAlo As New clsAlodine_planificacion
    objAlo.Imprimir_Etiquetas CLng(lista.ListItems(lista.selectedItem.Index).SubItems(7)), CLng(Index)
    Set objAlo = Nothing
End Sub

Private Sub cmdLimpiar_Click()
    txtFiltro(0) = ""
    txtFiltro(1) = ""
    txtFiltro(2) = ""
'    cmbProducto.BoundText = ""
'    cmbProducto.Text = ""
    txtanno = Year(Date)
    cmbclientes.limpiar
    cargar_lista
End Sub
Private Sub cambiar_Change()
    cargar_lista
End Sub

'Private Sub cmbproducto_Change()
'    cargar_lista
'End Sub

Private Sub cmdAnadir_Click()
    glote = 0
    frmAlodine_Lote.Show 1
    cargar_lista
End Sub

'Private Sub cmdDoc1_Click(Index As Integer)
'    On Error GoTo fallo
'    If lista.ListItems.Count > 0 Then
'        Dim documento As String
'        If Index = 0 Then
'            tipo = " CERT"
'        Else
'            tipo = ""
'        End If
'        documento = nombre_alodine(lista.ListItems(lista.selectedItem.Index).SubItems(7)) & tipo & ".pdf"
'        If Dir(documento) = "" Then
'            MsgBox "No se localiza el documento de alodine para imprimir.", vbCritical, App.Title
'        Else
'            Shell "rundll32.exe url.dll,FileProtocolHandler " & documento, vbMaximizedFocus
'        End If
'    End If
'    Exit Sub
'fallo:
'    MsgBox "Se ha producido un error al imprimir. Error : " + Err.Description, vbInformation, App.Title
'End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el Lote del producto : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oAlodine_Lote As New clsAlodine_lotes
            If oAlodine_Lote.Eliminar(lista.ListItems(lista.selectedItem.Index).SubItems(7)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

'Private Sub cmdEtiqueta_Click_old()
'    Dim oAlodine_planificacion As New clsAlodine_planificacion
'    If chkprev.value = Unchecked Then
'        If MsgBox("Va a generar " & oAlodine_planificacion.Numero_Etiquetas_Lote(CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(5))) & " etiquetas. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
'            Exit Sub
'        End If
'    End If
'    frmPosicionPegatina.Show 1
'    If pegatina <> 0 Then
'        On Error GoTo fallo
'        Dim i As Integer
'        ' Generamos los datos del listado
'        Dim rs As New ADODB.RecordSet
'        ' Recordset ficticio
'        Dim rs2 As New ADODB.RecordSet
'        rs2.Fields.Append "c1", adChar, 5, adFldUpdatable
'        rs2.Open
'        rs2.AddNew
'        rs2("c1") = Left(lista.ListItems(lista.SelectedItem.Index), 5)
'        rs2.Update
'        ' Comienzo
'        Dim Listado As New rptEtiquetaAlodine
'        ' Ocultar controles
'        For i = 1 To Listado.Sections("detalle").Controls.Count
'            Listado.Sections("detalle").Controls(i).Visible = False
'        Next
'        Dim num_bote As Integer
'        Set rs = oAlodine_planificacion.Listado_Etiquetas_Lote(CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(5)))
'        If rs.RecordCount > 0 Then
'            ' Generamos los datos del listado
'            num_bote = 0
'            Do
'repite:
'                With Listado.Sections("detalle")
'                    ' Logo
'                    If rs(6) = 0 Then ' Canagrosa
''                        Set .Controls(Trim("logo" & pegatina)).Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'                        .Controls(Trim("f" & pegatina)).Caption = ""
'                        .Controls(Trim("g" & pegatina)).Caption = "C/Wilbur y Orville Wright Nº 15, AERÓPOLIS, Tlf:954 115011"
''                        .Controls(Trim("g" & pegatina)).Caption = "Ed.Izpal Quinta Av. Nave I Pol.Ind.La Negrilla Tlf:954 25 88 78"
''                        .Controls("pie1").Caption = "C/Wilbur y Orville Wright Nº 15,P. T. AEROSPACIAL (AERÓPOLIS), 41309 La Rinconada (Sevilla) - Teléfonos +34 954 115011 Fax +34 954 115030 E-mail:consultoria@canagrosa.com"
'
'                    Else
''                        Set .Controls(Trim("logo" & pegatina)).Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "eads"))
''                        .Controls(Trim("f" & pegatina)).Caption = "Prep.: CANAGROSA (Ed.Izpal Quinta Av. Nave I Pol.Ind.La Negrilla.Sevilla. Tlf:954 25 88 78)"
'                        .Controls(Trim("f" & pegatina)).Caption = "Prep.: CANAGROSA (C/Wilbur y Orville Wright Nº 15, AERÓPOLIS, Tlf:954 115011)"
'                        .Controls(Trim("g" & pegatina)).Caption = ""
'                    End If
'                    ' Posicionar el logo en su posición original.
'                    .Controls(Trim("logo" & pegatina)).Width = 1755
'                    .Controls(Trim("logo" & pegatina)).Height = 675
'                    If pegatina Mod 2 <> 0 Then
'                        .Controls(Trim("logo" & pegatina)).Left = 1984
'                    Else
'                        .Controls(Trim("logo" & pegatina)).Left = 7650
'                    End If
'                    Select Case rs(8) ' Etiqueta
'                    Case 1 ' Canagrosa
'                        Set .Controls(Trim("logo" & pegatina)).Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'                    Case 2 ' EADS
'                        Set .Controls(Trim("logo" & pegatina)).Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "eads"))
'                    Case 3 ' Airbus
'                        Set .Controls(Trim("logo" & pegatina)).Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "airbus"))
'                        .Controls(Trim("logo" & pegatina)).Left = .Controls(Trim("logo" & pegatina)).Left - 964
'                        .Controls(Trim("logo" & pegatina)).Width = 2835
'                        .Controls(Trim("logo" & pegatina)).Height = 330
'                    End Select
'                    .Controls(Trim("a" & pegatina)).Caption = rs(0) ' Producto
'                    If rs(8) = 3 Then
'                        .Controls(Trim("b" & pegatina)).Caption = rs(1) & " Z23.401 (Cont. " & rs(5) & ")" ' Codigo Producto
'                    Else
'                        .Controls(Trim("b" & pegatina)).Caption = rs(1) & " (Cont. " & rs(5) & ")" ' Codigo Producto
'                    End If
'                    .Controls(Trim("c" & pegatina)).Caption = rs(2) ' Fecha fabricacion
'                    .Controls(Trim("d" & pegatina)).Caption = rs(3) & "/" & Year(rs(2)) ' Numero Lote
'                    .Controls(Trim("e" & pegatina)).Caption = rs(4) ' Caducidad
'
'                    .Controls(Trim("logo" & pegatina)).Visible = True
'                    .Controls(Trim("a" & pegatina)).Visible = True
'                    .Controls(Trim("b" & pegatina)).Visible = True
'                    .Controls(Trim("c" & pegatina)).Visible = True
'                    .Controls(Trim("d" & pegatina)).Visible = True
'                    .Controls(Trim("e" & pegatina)).Visible = True
'                    .Controls(Trim("f" & pegatina)).Visible = True
'                    .Controls(Trim("g" & pegatina)).Visible = True
'                    .Controls(Trim("aa" & pegatina)).Visible = True
'                    .Controls(Trim("bb" & pegatina)).Visible = True
'                    .Controls(Trim("cc" & pegatina)).Visible = True
'                    .Controls(Trim("dd" & pegatina)).Visible = True
'                    .Controls(Trim("ee" & pegatina)).Visible = True
'                    .Controls(Trim("ff" & pegatina)).Visible = True
'                End With
'                pegatina = pegatina + 1
'                num_bote = num_bote + 1
'                If num_bote = rs(7) Then
'                    rs.MoveNext
'                    num_bote = 0
'                End If
'            Loop Until rs.EOF Or pegatina = 13
'            Set Listado.DataSource = rs2
'            'If chkprev.value = Checked Then
'            '    Listado.Caption = "Etiquetas de Alodine"
'            '    Listado.WindowState = vbMaximized
'            '    Listado.Show 1
'            'Else
'                Listado.PrintReport
'                ' Ocultar controles
'                For i = 1 To Listado.Sections("detalle").Controls.Count
'                    Listado.Sections("detalle").Controls(i).Visible = False
'                Next
'                pegatina = 1
'                If rs.EOF = False Then
'                    GoTo repite
'                End If
'            'End If
'        End If
'        Set rs = Nothing
'    End If
'    If chkprev.value = Unchecked Then
'        Set Listado = Nothing
'    End If
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el documento.", vbCritical, Err.Description
'End Sub
'
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        glote = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        frmAlodine_Lote.Show 1
        actualizar_lista
        glote = 0
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fecha_f_Change()
    cargar_lista
End Sub

Private Sub fecha_i_Change()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
'    txtanno = Year(Date)
    
    fecha_i = "01/01/" & Year(Date)
    fecha_f = Date
    
    cabecera
    rellenar_clientes
'    cargar_combo cmbProducto, New clsAlodine
    cargar_lista
End Sub
Private Sub rellenar_clientes()
    Dim consulta As String
    consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
               "  FROM ALODINE_PLANIFICACION AP, CLIENTES C " & _
               " WHERE AP.CLIENTE_ID = C.ID_CLIENTE "
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbclientes
            .setCONN = conn
            .setQUERY = consulta
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "CLIENTES"
            .setDESCRIPCION = "Clientes"
            .setPK = "C.ID_CLIENTE"
            .setFILTRO = ""
            .setCAMPO = "C.NOMBRE"
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmClientes
        End With
    End If
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Producto", 3900, lvwColumnLeft
        .Add , , "Codigo", 1950, lvwColumnCenter
        .Add , , "Lote Componentes", 2200, lvwColumnCenter
        .Add , , "Nº Lote", 900, lvwColumnCenter
        .Add , , "Numero", 900, lvwColumnCenter
        .Add , , "F.Alta", 1100, lvwColumnCenter
        .Add , , "Caducidad", 1100, lvwColumnCenter
        .Add , , "ID", 800, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oAlodine_Lote As New clsAlodine_lotes
    lista.ListItems.Clear
'    Set rs = oAlodine_Lote.Listado(txtanno, txtfiltro(0), txtfiltro(1), cmbProducto.BoundText)
    Set rs = oAlodine_Lote.Listado(fecha_i, fecha_f, txtFiltro(0), txtFiltro(1), txtFiltro(2), IIf(cmbclientes.getTEXTO = "", 0, cmbclientes.getPK_SALIDA))
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(6)
             .SubItems(7) = rs(7)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oAlodine_Lote = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub
Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub
Public Sub actualizar_lista()
    Dim oAlodine_Lote As New clsAlodine_lotes
    With oAlodine_Lote
        If .Carga(glote) = True Then
            Dim oalodine As New clsAlodine
            oalodine.Carga (.getALODINE_ID)
            lista.ListItems(lista.selectedItem.Index).Text = oalodine.getPRODUCTO
            lista.ListItems(lista.selectedItem.Index).SubItems(1) = oalodine.getCODIGO
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = oalodine.getLOTE
'            lista.ListItems(lista.selectedItem.Index).SubItems(3) = oAlodine.getID_ALODINE
            lista.ListItems(lista.selectedItem.Index).SubItems(3) = oalodine.getMADRE
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = .getNUMERO_LOTE
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = .getFECHA_ALTA
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = .getFECHA_CADUCIDAD
        End If
    End With
    Set oalodine = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
