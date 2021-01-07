VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoParametrosTecnicos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Parámetros Técnicos de Equipos"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   Icon            =   "frmEquipoParametrosTecnicos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11025
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9915
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   810
      TabIndex        =   0
      Top             =   7050
      Width           =   9075
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6195
      Left            =   60
      TabIndex        =   5
      Top             =   765
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   10927
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
   Begin pryCombo.miCombo cmbNormas 
      Height          =   330
      Left            =   810
      TabIndex        =   10
      Top             =   7470
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   582
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Parámetros Técnicos para las Especificaciones de los Equipos de Ensayo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   420
      Width           =   6030
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10395
      Picture         =   "frmEquipoParametrosTecnicos.frx":08CA
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento de Parámetros Técnicos de Equipos"
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
      TabIndex        =   8
      Top             =   75
      Width           =   5370
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Norma"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   7560
      Width           =   465
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre"
      Height          =   195
      Index           =   6
      Left            =   105
      TabIndex        =   6
      Top             =   7140
      Width           =   555
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   11040
   End
End
Attribute VB_Name = "frmEquipoParametrosTecnicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
        Exit Sub
    End If
    If MsgBox("Va a insertar el Parámetro. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oEP As New clsEq_parametros_tecnicos
        With oEP
            .setNOMBRE = txtDatos(0)
            If cmbNormas.getTEXTO = "" Then
                .setNORMA_ID = 0
            Else
                .setNORMA_ID = cmbNormas.getPK_SALIDA
            End If
            .Insertar
        End With
        Set oEP = Nothing
        cargar_lista
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR el Parámetro. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oEP As New clsEq_parametros_tecnicos
            With oEP
                .Eliminar lista.ListItems(lista.SelectedItem.Index).Text
            End With
            cargar_lista
            Set oEP = Nothing
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
        Exit Sub
    End If
    If MsgBox("Va a modificar el proceso base. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oEP As New clsEq_parametros_tecnicos
        With oEP
            .setNOMBRE = txtDatos(0)
            If cmbNormas.getTEXTO = "" Then
                .setNORMA_ID = 0
            Else
                .setNORMA_ID = cmbNormas.getPK_SALIDA
            End If
            .Modificar lista.ListItems(lista.SelectedItem.Index).Text
            cargar_lista
            txtDatos(0) = ""
        End With
        Set oEP = Nothing
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    llenar_combo cmbNormas, New clsCa_normas, 0, frmCA_Normas, ""
    With lista.ColumnHeaders
        .Add , , "ID", 400, lvwColumnLeft
        .Add , , "Nombre", 5000, lvwColumnLeft
        .Add , , "Norma", 5000, lvwColumnLeft
        .Add , , "ID_NORMA", 1, lvwColumnLeft
    End With
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.RecordSet
    Dim oEP As New clsEq_parametros_tecnicos
    Set rs = oEP.Listado
    txtDatos(0) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
            .SubItems(1) = rs(1)
            If IsNull(rs(2)) Then
                .SubItems(2) = ""
                .SubItems(3) = "0"
            Else
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEP = Nothing
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        If lista.ListItems(lista.SelectedItem.Index).SubItems(3) = 0 Then
            cmbNormas.Limpiar
        Else
            cmbNormas.MostrarElemento lista.ListItems(lista.SelectedItem.Index).SubItems(3)
        End If
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
