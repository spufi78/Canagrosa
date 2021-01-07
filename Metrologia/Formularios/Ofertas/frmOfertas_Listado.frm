VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOfertas_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Ofertas"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
   Icon            =   "frmOfertas_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   13110
   Begin VB.CommandButton cmdDuplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   885
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdCorreo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correo"
      Height          =   885
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   885
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7995
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de búsqueda"
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
      Height          =   690
      Left            =   45
      TabIndex        =   12
      Top             =   360
      Width           =   13005
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   885
         TabIndex        =   0
         Top             =   240
         Width           =   885
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   6105
         TabIndex        =   3
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   4335
         TabIndex        =   2
         Top             =   240
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo cmbAgente 
         Height          =   315
         Left            =   7950
         TabIndex        =   4
         Top             =   270
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   3060
         TabIndex        =   6
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196610
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSDataListLib.DataCombo cmbEstado 
         Height          =   315
         Left            =   10770
         TabIndex        =   5
         Top             =   270
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Left            =   10200
         TabIndex        =   18
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   315
         Index           =   0
         Left            =   1935
         TabIndex        =   17
         Top             =   300
         Width           =   375
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   16
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obra"
         Height          =   195
         Index           =   1
         Left            =   5625
         TabIndex        =   15
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Agente"
         Height          =   195
         Left            =   7380
         TabIndex        =   14
         Top             =   330
         Width           =   510
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   3585
         TabIndex        =   13
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2420
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1225
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   11910
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7995
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7995
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6870
      Left            =   60
      TabIndex        =   11
      Top             =   1065
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   12118
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14609914
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Ofertas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   3
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   13035
   End
End
Attribute VB_Name = "frmOfertas_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCorreo_Click()
    If lista.ListItems.Count > 0 Then
        Dim oOferta As New clsOfertas
        oOferta.Correo lista.ListItems(lista.SelectedItem.Index).Text, False, True, 1
        Set oOferta = Nothing
    End If
End Sub

Private Sub cmdDuplicar_Click()
   On Error GoTo cmdDuplicar_Click_Error

    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar la oferta nº : " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & " ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim OFERTA As Long
            Dim oO As New clsOfertas
            Dim oOC As New clsOfertas
            ' Oferta
            oO.Carga (lista.ListItems(lista.SelectedItem.Index).Text)
            With oOC
                .setFECHA = Format(Date, "yyyy-mm-dd")
                .setCLIENTE_ID = oO.getCLIENTE_ID
                .setESTADO_ID = ENUM_OFERTAS_ESTADOS.OFERTAR_ESTADOS_PENDIENTE
                .setNIF = oO.getNIF
                .setNOMBRE = oO.getNOMBRE
                .setDIRECCION = oO.getDIRECCION
                .setPOBLACION = oO.getPOBLACION
                .setAA = oO.getAA
                .setTELEFONO = oO.getTELEFONO
                .setFAX = oO.getFAX
                .setEMAIL = oO.getEMAIL
                .setOBRA_DOMICILIO = oO.getOBRA_DOMICILIO
                .setOBRA_POBLACION = oO.getOBRA_POBLACION
                .setOBRA_TIPO = oO.getOBRA_TIPO
                .setFORMA_PAGO = oO.getFORMA_PAGO
                .setAGENTE_ID = oO.getAGENTE_ID
                .setHORA = Time
                OFERTA = .Insertar
            End With
            ' Detalle
            Dim rs As ADODB.Recordset
            Dim oOD As New clsOfertas_detalle
            Dim oODC As New clsOfertas_detalle
            Set rs = oOD.Listado(lista.ListItems(lista.SelectedItem.Index))
            Dim orden As Integer
            orden = 0
            If rs.RecordCount > 0 Then
                Do
                    With oODC
                        .setOFERTA_ID = OFERTA
                        .setORDEN = orden
                        .setMATERIAL_ID = rs(0)
                        .setPRECIO_FABRICA = rs(2)
                        .setPRECIO_OBRA = rs(3)
                        .Insertar
                        orden = orden + 1
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            MsgBox "Se ha duplicado correctamente la Oferta.", vbInformation + vbOKOnly, App.Title
            cargar_lista
            frmOfertas_Detalle.pk = OFERTA
            frmOfertas_Detalle.Show 1
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdDuplicar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDuplicar_Click of Formulario frmOfertas_Listado"
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        Dim oOferta As New clsOfertas
        oOferta.imprimir lista.ListItems(lista.SelectedItem.Index).Text, False, True, 1, ""
        Set oOferta = Nothing
    End If
End Sub

Private Sub cambiar_Change()
    cargar_lista
End Sub
Private Sub cmbAgente_Change()
    cargar_lista
End Sub

Private Sub cmbEstado_Change()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmOfertas_Detalle.pk = 0
    frmOfertas_Detalle.Show 1
    cargar_lista
    lista.SetFocus
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la Oferta " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim oOferta As New clsOfertas
            If oOferta.Eliminar(lista.ListItems(lista.SelectedItem.Index).Text) = True Then
                cargar_lista
            End If
            Set oOferta = Nothing
        End If
        lista.SetFocus
    End If
End Sub

Private Sub cmdModificar_Click()
    If USUARIO.getPER_3 = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        frmOfertas_Detalle.pk = lista.ListItems(lista.SelectedItem.Index)
        frmOfertas_Detalle.Show 1
        actualizar_lista
        lista.SetFocus
    End If
End Sub



Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 100
    Me.Top = 100
    txtanno = Year(Date)
    cambiar.Max = Year(Date)
    cabecera
    cargar_combos
    cargar_lista
    permisos
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oOferta As New clsOfertas
    Set rs = oOferta.Listado(txtDatos(0), txtDatos(1), txtDatos(2), cmbAgente.BoundText, txtanno, cmbEstado.BoundText)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0)) ' ID
            .SubItems(1) = rs(1) ' NUMERO
            .SubItems(2) = Format(rs(2), "dd-mm-yyyy") ' FECHA
            .SubItems(3) = rs(3) ' CLIENTE
            .SubItems(4) = rs(4) ' OBRA
            .SubItems(5) = rs(5) ' AGENTE
            .SubItems(6) = rs(6) ' ESTADO
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oOferta = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
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
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index) <> "" Then
          cmdmodificar.Enabled = True
          cmdeliminar.Enabled = True
        End If
        permisos
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim oOferta As New clsOfertas
    Dim rs As ADODB.Recordset
    Set rs = oOferta.Listado_ID(lista.ListItems(lista.SelectedItem.Index).Text)
    If rs.RecordCount > 0 Then
        With lista.ListItems(lista.SelectedItem.Index) ' ID
         .SubItems(1) = rs(1) ' NUMERO
         .SubItems(2) = Format(rs(2), "dd-mm-yyyy") ' FECHA
         .SubItems(3) = rs(3) ' CLIENTE
         .SubItems(4) = rs(4) ' OBRA
         .SubItems(5) = rs(5) ' AGENTE
         .SubItems(6) = rs(6) ' ESTADO
        End With
    End If
    Set oOferta = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       cmdModificar_Click
    End If
End Sub
Public Sub permisos()
    If USUARIO.getPER_1 = 0 Then
'        cmdImprimir.Enabled = False
    End If
    If USUARIO.getPER_2 = 0 Then
        cmdanadir.Enabled = False
    End If
    If USUARIO.getPER_3 = 0 Then
        cmdmodificar.Enabled = False
    End If
    If USUARIO.getPER_4 = 0 Then
        cmdeliminar.Enabled = False
    End If
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Número", 800, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Cliente", 3500, lvwColumnLeft
        .Add , , "Obra", 3500, lvwColumnLeft
        .Add , , "Agente", 2000, lvwColumnCenter
        .Add , , "Estado", 1800, lvwColumnCenter
    End With
End Sub

Private Sub cargar_combos()
    Cargar_Combo cmbAgente, New clsComercial
    Dim oD As New clsDecodificadora
    oD.Cargar_Combo cmbEstado, DECODIFICADORA.D_OFERTAS_ESTADOS
End Sub
