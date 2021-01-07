VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContabilidad_Movimientos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Movimientos Contables"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
   Icon            =   "frmContabilidad_Movimientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   13110
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
      Height          =   870
      Left            =   60
      TabIndex        =   2
      Top             =   375
      Width           =   13005
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   10890
         TabIndex        =   11
         Top             =   360
         Width           =   1410
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5850
         TabIndex        =   10
         Top             =   360
         Width           =   1245
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   8280
         TabIndex        =   9
         Top             =   360
         Width           =   1425
      End
      Begin VB.CheckBox chkFecha 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   285
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1530
         TabIndex        =   5
         Top             =   360
         Width           =   1350
         _ExtentX        =   2381
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
         CalendarTitleBackColor=   12632256
         Format          =   50593793
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3720
         TabIndex        =   6
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
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
         CalendarTitleBackColor=   12632256
         Format          =   50593793
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta"
         Height          =   195
         Index           =   1
         Left            =   5220
         TabIndex        =   14
         Top             =   420
         Width           =   510
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   9900
         TabIndex        =   13
         Top             =   420
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contrapartida"
         Height          =   195
         Index           =   2
         Left            =   7245
         TabIndex        =   12
         Top             =   420
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Desde"
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   8
         Top             =   420
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   3090
         TabIndex        =   7
         Top             =   420
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   11910
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7995
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6645
      Left            =   60
      TabIndex        =   1
      Top             =   1290
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   11721
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
      Caption         =   "Listado de Movimientos Contables"
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
      TabIndex        =   3
      Top             =   0
      Width           =   13035
   End
End
Attribute VB_Name = "frmContabilidad_Movimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        fdesde.Enabled = True
        fhasta.Enabled = True
    Else
        fdesde.Enabled = False
        fhasta.Enabled = False
    End If
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub fdesde_LostFocus()
    cargar_lista
End Sub
Private Sub fhasta_LostFocus()
    cargar_lista
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
    With lista.ColumnHeaders
        .Add , , "Apunte", 700, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Documento", 900, lvwColumnLeft
        .Add , , "Cuenta", 2500, lvwColumnLeft
        .Add , , "Descripción", 3500, lvwColumnLeft
        .Add , , "Importe", 1000, lvwColumnRight
        .Add , , "Tipo", 400, lvwColumnCenter
        .Add , , "Contrapartida", 2500, lvwColumnLeft
    End With
    fdesde = Date - 31
    fhasta = Date
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim oContabilidad As New clsContabilidad
   On Error GoTo cargar_lista_Error

    If Not oContabilidad.Conectar Then
        Exit Sub
    End If
    
    Dim rs As New ADODB.Recordset
    Dim oMovimiento As New clsContabilidad_Movimientos
        
    Set rs = oMovimiento.Listado(chkFecha.Value, fdesde, fhasta, txtDatos(0), txtDatos(1), txtDatos(2))
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "000"))
            .SubItems(1) = rs(1) ' FECHA
            .SubItems(2) = rs(2) ' DOCUMENTO
            .SubItems(3) = rs(3) & " (" & rs(4) & ")" ' CUENTA
            .SubItems(4) = rs(5) ' DESCRIPCION
            .SubItems(5) = moneda(rs(6)) ' IMPORTE
            .SubItems(6) = rs(7) ' TIPO
            .SubItems(7) = rs(8) & " (" & rs(9) & ")" ' CONTRAPARTIDA
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    
    
    oContabilidad.DesConectar

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmContabilidad_Movimientos"
    
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
