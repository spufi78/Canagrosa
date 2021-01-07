VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCE_Materiales_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Materiales"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmCE_Materiales_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10395
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
      Height          =   825
      Left            =   45
      TabIndex        =   9
      Top             =   585
      Width           =   10275
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   6030
         MaxLength       =   255
         TabIndex        =   1
         Top             =   360
         Width           =   3120
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   900
         MaxLength       =   255
         TabIndex        =   0
         Top             =   360
         Width           =   3120
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Material"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   405
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Criterio"
         Height          =   195
         Index           =   3
         Left            =   5355
         TabIndex        =   10
         Top             =   405
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6525
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2270
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6525
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6525
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6525
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6525
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5040
      Left            =   45
      TabIndex        =   2
      Top             =   1440
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   8890
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
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Materiales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   180
      TabIndex        =   8
      Top             =   90
      Width           =   9435
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9765
      Picture         =   "frmCE_Materiales_Listado.frx":08CA
      Top             =   45
      Width           =   480
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
Attribute VB_Name = "frmCE_Materiales_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    frmCE_Materiales.PK = 0
    frmCE_Materiales.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdduplicar_Click()
   On Error GoTo cmdduplicar_Click_Error

    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar el MATERIAL, ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oMAT As Long
            Dim oCE_Mat As New clsCe_banos_materiales
            Dim oCe_Mat_Copia As New clsCe_banos_materiales
            If oCE_Mat.Carga(lista.ListItems(lista.SelectedItem.Index).Text) Then
                ' Tipos de Ensayos
                With oCe_Mat_Copia
                    .setMATERIAL = oCE_Mat.getMATERIAL & " (Duplicado)"
                    .setMINIMO = oCE_Mat.getMINIMO
                    .setMAXIMO = oCE_Mat.getMAXIMO
                    .setMINIMO_TEXTO = oCE_Mat.getMINIMO_TEXTO
                    .setMAXIMO_TEXTO = oCE_Mat.getMAXIMO_TEXTO
                    .setCRITERIO = oCE_Mat.getCRITERIO
                    oMAT = .Insertar
                End With
                MsgBox "Se ha generado el Material correctamente.", vbInformation + vbOKOnly, App.Title
                cargar_lista
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdduplicar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdduplicar_Click of Formulario frmCE_Materiales_Listado"
    
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el MATERIAL : " & lista.ListItems(lista.SelectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oCE_Mat As New clsCe_banos_materiales
            If oCE_Mat.Eliminar(lista.ListItems(lista.SelectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmCE_Materiales.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmCE_Materiales.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.Top = 50
    cabecera
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Material", 5000, lvwColumnLeft
        .Add , , "Criterio", 5000, lvwColumnLeft
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.RecordSet
    Dim oCE_Material As New clsCe_banos_materiales
    lista.ListItems.Clear
    Set rs = oCE_Material.Listado(txtDatos(0), txtDatos(1))
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID_MATERIAL"))
             .SubItems(1) = rs("MATERIAL")
             .SubItems(2) = rs("CRITERIO")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oCE_Material = Nothing
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

Public Sub actualizar_lista()
    Dim rs As ADODB.RecordSet
    Dim oCE_Material As New clsCe_banos_materiales
    If oCE_Material.Carga(CLng(lista.ListItems(lista.SelectedItem.Index).Text)) Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = oCE_Material.getMATERIAL
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = oCE_Material.getCRITERIO
    End If
    Set oCE_Material = Nothing
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub
