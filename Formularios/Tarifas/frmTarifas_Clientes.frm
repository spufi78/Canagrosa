VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Begin VB.Form frmTarifas_Clientes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tarifas"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9480
   ClipControls    =   0   'False
   Icon            =   "frmTarifas_Clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAsignar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignar la tarifa seleccionada a los clientes de la lista"
      Height          =   375
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8190
      Width           =   5550
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8340
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7725
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6885
      Left            =   0
      TabIndex        =   1
      Top             =   780
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   12144
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
   Begin pryCombo.miCombo cmbtarifa 
      Height          =   330
      Left            =   585
      TabIndex        =   4
      Top             =   7830
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   582
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tarifa"
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   7875
      Width           =   465
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Existen : "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   420
      Width           =   645
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clientes con la tarifa : "
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
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   2310
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8805
      Picture         =   "frmTarifas_Clientes.frx":08CA
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   9465
   End
End
Attribute VB_Name = "frmTarifas_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmdAsignar_Click()
    If lista.ListItems.Count > 0 Then
        If cmbtarifa.getTEXTO <> "" Then
            If MsgBox("¿Esta seguro de asignar la tarifa " & cmbtarifa.getTEXTO & " a los clientes de la lista?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Dim oCliente As New clsCliente
                Dim i As Integer
                For i = 1 To lista.ListItems.Count
                    oCliente.Modificar_Tarifa lista.ListItems(i).Text, cmbtarifa.getPK_SALIDA
                Next
                MsgBox "Tarifas asignadas correctamente.", vbInformation, App.Title
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 200
    Me.Left = 200
    cargar_botones Me
    llenar_combo cmbtarifa, New clsTarifas, 0, Me, ""
    Dim oTarifa As New clsTarifas
    oTarifa.Carga PK
    lbltitulo = lbltitulo & oTarifa.getNOMBRE
    Me.Caption = lbltitulo
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Cliente", lista.Width - 250, lvwColumnLeft
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.RecordSet
    Dim oTarifa As New clsTarifas
    Set rs = oTarifa.Listado_Clientes(PK)
    lista.ListItems.Clear
    lblsubtitulo = lblsubtitulo & rs.RecordCount & " Clientes"
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
            .SubItems(1) = rs(1)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
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
    If lista.ListItems.Count > 0 Then
        frmClientes.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmClientes.Show 1
    End If
End Sub
