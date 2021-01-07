VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPinturasRecepcionAdministrativaListado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción Administrativa de PINTURAS"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmPinturasRecepcionAdministrativaListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   11670
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
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
      Height          =   1140
      Left            =   45
      TabIndex        =   9
      Top             =   585
      Width           =   11580
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1050
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   900
         TabIndex        =   0
         Top             =   270
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbEstado 
         Height          =   330
         Left            =   6840
         TabIndex        =   1
         Top             =   270
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   900
         TabIndex        =   13
         Top             =   675
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   2835
         TabIndex        =   14
         Top             =   675
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   765
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   2340
         TabIndex        =   15
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   0
         Left            =   6300
         TabIndex        =   12
         Top             =   315
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clientes"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7605
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5760
      Left            =   45
      TabIndex        =   8
      Top             =   1755
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   10160
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
   Begin VB.CommandButton cmdRecepcionar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recepcionar"
      Height          =   870
      Left            =   9450
      Picture         =   "frmPinturasRecepcionAdministrativaListado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7605
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recepción Administrativa de PINTURAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   180
      TabIndex        =   11
      Top             =   90
      Width           =   9435
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   11790
   End
End
Attribute VB_Name = "frmPinturasRecepcionAdministrativaListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkactiva_Click()
    cargar_lista
End Sub

Private Sub cmbCE_change()
    cargar_lista
End Sub

Private Sub cmbTA_change()
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    fhasta = Date
    fdesde = Date - 365
    cmbclientes.Limpiar
    cmbEstado.Limpiar
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmPinturasRecepcionAdministrativa.PK = 0
    frmPinturasRecepcionAdministrativa.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub


Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar la Recepción Administrativa, ¿esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oPR As New clsPinturas_radmin
            If oPR.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmPinturasRecepcionAdministrativa.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmPinturasRecepcionAdministrativa.Show 1
        actualizar_lista
    End If
End Sub

Private Sub cmdRecepcionar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oPR As New clsPinturas_radmin
    If oPR.Carga(lista.ListItems(lista.selectedItem.Index).Text) = True Then
        If oPR.getESTADO_ID <> PINTURAS_ESTADOS.PINTURAS_PENDIENTE Then
            MsgBox "La recepción no se encuentra en un estado para poder Recepcionar.", vbCritical, App.Title
        Else
            frmPinturasRecepcion.PK = lista.ListItems(lista.selectedItem.Index).Text
            frmPinturasRecepcion.Show 1
            actualizar_lista
        End If
    End If
    Set oPR = Nothing
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    fhasta = Date
    fdesde = Date - 365
    cargar_lista
    cargar_combos
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_RADMIN", 1, lvwColumnLeft
        .Add , , "Cliente", 5000, lvwColumnLeft
        .Add , , "Fecha", 1300, lvwColumnCenter
        .Add , , "Estado", 1500, lvwColumnCenter
        .Add , , "Usuario", 1200, lvwColumnCenter
        .Add , , "F.Creación", 1800, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPR As New clsPinturas_radmin
    lista.ListItems.Clear
    Dim cCliente As Long
    Dim cEstado As Long
    cCliente = 0
    cEstado = 0
    If cmbclientes.getTEXTO <> "" Then
        cCliente = cmbclientes.getPK_SALIDA
    End If
    If cmbEstado.getTEXTO <> "" Then
        cEstado = cmbEstado.getPK_SALIDA
    End If
    Set rs = oPR.Listado(fdesde, fhasta, cCliente, cEstado)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0)) 'ID_RADMIN
             .SubItems(1) = rs(1) ' CLIENTE
             .SubItems(2) = Format(rs(2), "dd-mm-yyyy") ' FECHA
             .SubItems(3) = rs(3)  ' ESTADO
             .SubItems(4) = rs(4)  ' USUARIOS
             .SubItems(5) = rs(5)  ' FS
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPR = Nothing
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
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub actualizar_lista()
    Dim oPR As New clsPinturas_radmin
    Dim rs As ADODB.Recordset
    Set rs = oPR.ListadoID(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount <> 0 Then
        With lista.ListItems(lista.selectedItem.Index)
         .SubItems(1) = rs(1) ' CLIENTE
         .SubItems(2) = Format(rs(2), "dd-mm-yyyy") ' FECHA
         .SubItems(3) = rs(3)  ' ESTADO
         .SubItems(4) = rs(4)  ' USUARIOS
         .SubItems(5) = rs(5)  ' FS
        End With
    End If
    Set oPR = Nothing
End Sub

Private Sub txtDatos_Change()
    cargar_lista
End Sub

Private Sub cargar_combos()
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbEstado, DECODIFICADORA.PINTURAS_ESTADOS
    Set oDeco = Nothing
End Sub
