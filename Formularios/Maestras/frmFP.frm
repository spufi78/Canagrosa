VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFP 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Formas de Pago"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   Icon            =   "frmFP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6570
      Width           =   1080
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
      Height          =   360
      Index           =   2
      Left            =   5400
      TabIndex        =   9
      Top             =   6165
      Width           =   4200
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6570
      Width           =   1050
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
      Height          =   360
      Index           =   1
      Left            =   4275
      TabIndex        =   1
      Top             =   6165
      Width           =   1080
   End
   Begin VB.TextBox txtDatos 
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
      Height          =   360
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   6165
      Width           =   4200
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6570
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6570
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5445
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   9604
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "C.C.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5415
      TabIndex        =   10
      Top             =   5940
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   7
      Top             =   5940
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   5940
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mantenimiento de Formas de Pago"
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
      Index           =   3
      Left            =   45
      TabIndex        =   5
      Top             =   15
      Width           =   9555
   End
End
Attribute VB_Name = "frmFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "El nombre de la FP no puede estar en blanco.", vbCritical, App.Title
        txtDatos(0).SetFocus
    ElseIf txtDatos(1).Text = "" Then
        MsgBox "Los dias no pueden estar en blanco.", vbCritical, App.Title
        txtDatos(1).SetFocus
    Else
        If MsgBox("Va a insertar la Forma de Pago. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oFP As New clsFP
            With oFP
                .setNOMBRE = txtDatos(0)
                .setDIAS = txtDatos(1)
                .setCCC = txtDatos(2)
                .Insertar
            End With
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la Forma de Pago. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oFP As New clsFP
            oFP.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If txtDatos(0).Text = "" Then
        MsgBox "El nombre de la FP no puede estar en blanco.", vbCritical, App.Title
        txtDatos(0).SetFocus
    ElseIf txtDatos(1).Text = "" Then
        MsgBox "Los dias no pueden estar en blanco.", vbCritical, App.Title
        txtDatos(1).SetFocus
    Else
        If MsgBox("Va a modificar la Forma de Pago. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oFP As New clsFP
            With oFP
                .setNOMBRE = txtDatos(0)
                .setDIAS = txtDatos(1)
                .setCCC = txtDatos(2)
                .Modificar lista.ListItems(lista.selectedItem.Index).Text
            End With
            cargar_lista
        End If
    End If

End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "ID", 500, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Descripción", 3650, lvwColumnLeft)
        .Tag = "Caducidad"
    End With
    With lista.ColumnHeaders.Add(, , "Dias", 1250, lvwColumnCenter)
        .Tag = "Dias"
    End With
    With lista.ColumnHeaders.Add(, , "C.C.C.", 4000, lvwColumnCenter)
        .Tag = "C.C.C."
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim RS As New ADODB.Recordset
    Dim oFP As New clsFP
    Set RS = oFP.Listado
    txtDatos(0) = ""
    txtDatos(1) = ""
    txtDatos(2) = ""
    lista.ListItems.Clear
    If RS.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(RS("id_fp"), "000"))
            .SubItems(1) = RS("nombre")
            .SubItems(2) = RS("dias")
            .SubItems(3) = RS("ccc")
           End With
           RS.MoveNext
        Loop Until RS.EOF
    End If
    Set oFP = Nothing
    Set RS = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtDatos(1) = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        txtDatos(2) = lista.ListItems(lista.selectedItem.Index).SubItems(3)
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

