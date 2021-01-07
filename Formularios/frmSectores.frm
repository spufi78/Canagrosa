VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSectores 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Sectores"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "frmSectores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   9075
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
      Left            =   1230
      TabIndex        =   1
      Top             =   7200
      Width           =   2235
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7740
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7740
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7740
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7740
      Width           =   1050
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
      Left            =   1230
      TabIndex        =   0
      Top             =   6780
      Width           =   7755
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6375
      Left            =   60
      TabIndex        =   6
      Top             =   360
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   11245
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
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Departamento"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   9
      Top             =   7245
      Width           =   1005
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   8
      Top             =   6840
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento Sectores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   60
      TabIndex        =   7
      Top             =   15
      Width           =   8925
   End
End
Attribute VB_Name = "frmSectores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If txtDatos(1).Text = "" Then
            MsgBox "El departamento no puede estar en blanco.", vbCritical, App.Title
        Else
            If MsgBox("Va a insertar el sector. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                Dim olinea As New clsSectores
                With olinea
                    .setNOMBRE = txtDatos(0)
                    .setDEPARTAMENTO = txtDatos(1)
                    .Insertar
                End With
                txtDatos(0) = ""
                txtDatos(1) = ""
                cargar_lista
            End If
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR el sector. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim olinea As New clsSectores
            olinea.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If txtDatos(1).Text = "" Then
            MsgBox "El departamento no puede estar en blanco.", vbCritical, App.Title
        Else
            If MsgBox("Va a modificar el sector. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                Dim osec As New clsSectores
                osec.setNOMBRE = txtDatos(0)
                osec.setDEPARTAMENTO = txtDatos(1)
                osec.Modificar (lista.ListItems(lista.selectedItem.Index).Text)
                cargar_lista
                txtDatos(0) = ""
                txtDatos(1) = ""
            End If
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 200
    Me.top = 200
    With lista
        .ColumnHeaders.Add , , "ID", 1, lvwColumnLeft
        .ColumnHeaders.Add , , "Nombre", 6800, lvwColumnLeft
        .ColumnHeaders.Add , , "Departamento", 1500, lvwColumnCenter
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oSectores As New clsSectores
    Set rs = oSectores.Listado
    txtDatos(0) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("id_sector"))
            .SubItems(1) = rs("nombre")
            .SubItems(2) = rs("departamento")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSectores = Nothing
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
        txtDatos(0).Text = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtDatos(1).Text = lista.ListItems(lista.selectedItem.Index).SubItems(2)
    End If
End Sub

