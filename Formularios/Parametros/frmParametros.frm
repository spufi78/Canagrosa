VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParametros 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Parámetros del sistema"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   Icon            =   "frmParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   12000
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   30
      TabIndex        =   7
      Top             =   7020
      Width           =   11925
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   3
         Left            =   10050
         TabIndex        =   11
         Top             =   210
         Width           =   1785
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   2
         Left            =   6060
         TabIndex        =   10
         Top             =   210
         Width           =   3975
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   1
         Left            =   990
         TabIndex        =   9
         Top             =   210
         Width           =   5055
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   210
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7710
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7710
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7710
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7710
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6390
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   11271
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
      Caption         =   "Parámetros del sistema"
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
      TabIndex        =   6
      Top             =   45
      Width           =   2475
   End
   Begin VB.Image imagen 
      Height          =   420
      Left            =   11535
      Picture         =   "frmParametros.frx":08CA
      Top             =   15
      Width           =   420
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Identificación de Parámetros del sistema"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   330
      Width           =   2835
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   12030
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "El ID no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a insertar el Parámetro. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oParametro  As New clsParametros
            With oParametro
                .setID_PARAMETRO = txtDatos(0)
                .setDESCRIPCION = txtDatos(1)
                .setVALOR = txtDatos(2)
                .setUSUARIO = txtDatos(3)
                .Insertar
            End With
            cargar_lista
            txtDatos(0) = ""
            txtDatos(1) = ""
            txtDatos(2) = ""
            txtDatos(3) = ""
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR el parámetro. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oParametro  As New clsParametros
            oParametro.Eliminar lista.ListItems(lista.selectedItem.Index).Text, lista.ListItems(lista.selectedItem.Index).SubItems(3)
            cargar_lista
            txtDatos(0) = ""
            txtDatos(1) = ""
            txtDatos(2) = ""
            txtDatos(3) = ""
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "El ID no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar el parámetro. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oParametro  As New clsParametros
            With oParametro
                .setDESCRIPCION = txtDatos(1)
                .setVALOR = txtDatos(2)
                .setUSUARIO = txtDatos(3)
                .Modificar lista.ListItems(lista.selectedItem.Index).Text, lista.ListItems(lista.selectedItem.Index).SubItems(3)
            End With
            cargar_lista
            txtDatos(0) = ""
            txtDatos(1) = ""
            txtDatos(2) = ""
            txtDatos(3) = ""
        End If
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders
        .Add , , "ID", 900, lvwColumnLeft
        .Add , , "Descripcion", 4900, lvwColumnLeft
        .Add , , "Valor", 3900, lvwColumnCenter
        .Add , , "Usuario", 1800, lvwColumnCenter
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim RS As ADODB.Recordset
    Dim oParametros As New clsParametros
    Set RS = oParametros.Listado
    txtDatos(0) = ""
    lista.ListItems.Clear
    If RS.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , RS(0))
            .SubItems(1) = RS(1)
            .SubItems(2) = RS(2)
            .SubItems(3) = RS(3)
           End With
           RS.MoveNext
        Loop Until RS.EOF
    End If
    Set oParametros = Nothing
    Set RS = Nothing
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
        txtDatos(0).Text = lista.ListItems(lista.selectedItem.Index).Text
        txtDatos(1).Text = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtDatos(2).Text = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        txtDatos(3).Text = lista.ListItems(lista.selectedItem.Index).SubItems(3)
    End If
End Sub

