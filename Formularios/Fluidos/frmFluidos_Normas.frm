VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#27.0#0"; "miCombo.ocx"
Begin VB.Form frmFluidos_Normas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Normas de Fluidos"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmFluidos_Normas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modificar Descripción de la norma"
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   45
      TabIndex        =   6
      Top             =   6885
      Width           =   5100
      Begin VB.CommandButton Command1 
         Caption         =   "Modificar"
         Height          =   285
         Left            =   4050
         TabIndex        =   8
         Top             =   315
         Width           =   960
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   315
         Width           =   3900
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccione la norma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   45
      TabIndex        =   4
      Top             =   945
      Width           =   6615
      Begin pryCombo.miCombo cmbnorma 
         Height          =   345
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   609
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   5625
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6885
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5040
      Left            =   45
      TabIndex        =   0
      Top             =   1800
      Width           =   6615
      _ExtentX        =   11668
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
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Normas de Fluidos"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1980
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   6120
      Picture         =   "frmFluidos_Normas.frx":1272
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Muestra las normas existentes de fluidos y sus valores"
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   90
      TabIndex        =   2
      Top             =   420
      Width           =   4830
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   870
      Left            =   0
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "frmFluidos_Normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbnorma_change()
    cargar_lista (cmbnorma.getPK_SALIDA)
    If cmbnorma.getPK_SALIDA <> 0 Then
        txtDatos(0) = cmbnorma.getTEXTO
    Else
        txtDatos(0) = ""
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If txtDatos(0) <> "" Then
        If cmbnorma.getPK_SALIDA <> 0 Then
            Dim oFN As New clsFluidos_normas
            oFN.setNOMBRE = txtDatos(0)
            oFN.Modificar cmbnorma.getPK_SALIDA
            txtDatos(0) = ""
            MsgBox "La norma se ha modificado correctamente.", vbInformation, App.Title
            cmbnorma.Limpiar
            llenar_combo cmbnorma, New clsFluidos_normas, 0, frmClientes, ""
            cargar_lista (cmbnorma.getPK_SALIDA)
        End If
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    llenar_combo cmbnorma, New clsFluidos_normas, 0, frmClientes, ""
    cargar_lista (0)
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "GRADO", 1000, lvwColumnLeft
        .Add , , "5 a 15", 1000, lvwColumnCenter
        .Add , , "16 a 25", 1000, lvwColumnCenter
        .Add , , "26 a 50", 1000, lvwColumnCenter
        .Add , , "51 a 100", 1000, lvwColumnCenter
        .Add , , "> 100", 1000, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista(NORMA As Integer)
    Dim rs As ADODB.RecordSet
    Dim oFN As New clsFluidos_normas_valores
    lista.ListItems.Clear
    Set rs = oFN.Listado(NORMA)
    If rs.RecordCount <> 0 Then
        Do
            If rs("RANGO") <> rango_ant Then
                lista.ListItems.Add , , rs("RANGO")
                rango_ant = rs("RANGO")
            End If
            lista.ListItems(lista.ListItems.Count).SubItems(rs("tamano")) = rs("valor")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oFN = Nothing
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
