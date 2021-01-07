VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSoluciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Soluciones"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmSoluciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   9210
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
      Height          =   735
      Left            =   45
      TabIndex        =   11
      Top             =   360
      Width           =   9105
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1035
         TabIndex        =   12
         Top             =   270
         Width           =   2265
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.CheckBox chkFD 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Factura por Determinaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1080
      TabIndex        =   10
      Top             =   7425
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7875
      Width           =   1050
   End
   Begin VB.CommandButton cmdQuien 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¿Donde?"
      Height          =   870
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7875
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7875
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7875
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7875
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8145
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7875
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   6990
      Width           =   8115
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5835
      Left            =   60
      TabIndex        =   1
      Top             =   1125
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10292
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Solución"
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
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   7065
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mantenimiento Soluciones"
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
      Height          =   315
      Index           =   3
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   9120
   End
End
Attribute VB_Name = "frmSoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        With frmReport
            .iniciar
            .informe = "\Banos\rptsoluciones"
            .criterio = ""
            .imprimir = False
            .generar
            .Visible = True
        End With
    End If
End Sub
Private Sub cmdQuien_Click()
    If lista.ListItems.Count > 0 Then
        frmTD_Donde.solucion = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        frmTD_Donde.Show 1
    End If
End Sub
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a insertar el Proceso Base. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim opb As New clsSoluciones
            opb.setNOMBRE = txtDatos(0)
            opb.setFACTURA_DETERMINACIONES = chkFD.value
            opb.Insertar
            cargar_lista
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la solución. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim osol As New clsSoluciones
            osol.Eliminar (lista.ListItems(lista.selectedItem.Index).SubItems(1))
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar la solución. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim osol As New clsSoluciones
            osol.setNOMBRE = txtDatos(0)
            osol.setFACTURA_DETERMINACIONES = chkFD.value
            osol.Modificar (lista.ListItems(lista.selectedItem.Index).SubItems(1))
            cargar_lista
            txtDatos(0) = ""
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "Nombre", 7450, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 600, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Fac.Det.", 1, lvwColumnCenter)
        .Tag = "Fac.Det."
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.RecordSet
    Dim oSoluciones As New clsSoluciones
    Set rs = oSoluciones.Listado(txtfiltro)
    txtDatos(0) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("nombre"))
            .SubItems(1) = Format(rs("id_Solucion"), "0000")
            If rs("factura_determinaciones") = 0 Then
                .SubItems(2) = "No"
            Else
                .SubItems(2) = "Si"
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSoluciones = Nothing
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

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.selectedItem.Index).Text
        If lista.ListItems(lista.selectedItem.Index).SubItems(2) = "No" Then
            chkFD.value = 0
        Else
            chkFD.value = 1
        End If
    End If
End Sub

Private Sub txtfiltro_Change()
    cargar_lista
End Sub
