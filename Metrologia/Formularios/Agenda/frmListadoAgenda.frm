VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoAgenda 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   8595
   ClientLeft      =   135
   ClientTop       =   1425
   ClientWidth     =   11700
   Icon            =   "frmListadoAgenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11700
   Begin VB.TextBox txttexto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6300
      TabIndex        =   0
      Top             =   7950
      Width           =   2715
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7560
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   13335
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
      BackColor       =   &H00C0C0C0&
      Caption         =   "Introduzca nombre a buscar"
      Height          =   255
      Left            =   4020
      TabIndex        =   6
      Top             =   7980
      Width           =   2235
   End
End
Attribute VB_Name = "frmListadoAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    gAgenda = 0
    frmAgenda.Show 1
    If lista.ListItems.Count > 0 Then
       buscar_agenda
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR de la agenda " & lista.ListItems(lista.SelectedItem.Index) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim oagenda As New clsAgenda
            oagenda.setID_AGENDA = CInt(lista.ListItems(lista.SelectedItem.Index).SubItems(4))
            If oagenda.Eliminar = True Then
                If lista.ListItems.Count > 0 Then
                    buscar_agenda
                Else
                    lista.ListItems.Clear
                End If
            End If
            Set oagenda = Nothing
        End If
        lista.SetFocus
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        gAgenda = lista.ListItems(lista.SelectedItem.Index).SubItems(4)
        frmAgenda.Show 1
        If lista.ListItems.Count > 0 Then
           buscar_agenda
        End If
    End If
End Sub
Public Sub buscar_agenda()
    lista.ListItems.Clear
'    If txttexto <> "" Then
        Dim oagenda As New clsAgenda
        Dim RS As ADODB.Recordset
        Set RS = oagenda.Listado_por_letra(UCase(txttexto))
        If RS.RecordCount > 0 Then
            Do
               With lista.ListItems.Add(, , RS(0))
                .SubItems(1) = RS(1)
                .SubItems(2) = RS(2)
                .SubItems(3) = RS(3)
                .SubItems(4) = RS(4)
               End With
               RS.MoveNext
            Loop Until RS.EOF
        End If
'    End If
End Sub

Private Sub Form_Activate()
    buscar_agenda
    txttexto.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.Top = 50
    cabecera
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Nombre", 5300, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Teléfono", 2000, lvwColumnCenter)
        .Tag = "Teléfono"
    End With
    With lista.ColumnHeaders.Add(, , "Móvil", 2000, lvwColumnCenter)
        .Tag = "Móvil"
    End With
    With lista.ColumnHeaders.Add(, , "Fax", 2000, lvwColumnCenter)
        .Tag = "Fax"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
End Sub
Private Sub txttexto_Change()
    buscar_agenda
End Sub
