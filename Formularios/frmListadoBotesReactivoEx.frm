VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoBotesReactivoEx 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Botes de Reactivos Externos"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13575
   Icon            =   "frmListadoBotesReactivoEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13575
   Begin VB.CommandButton cmdmanual 
      Caption         =   "Código Manual"
      Height          =   645
      Left            =   9720
      TabIndex        =   39
      Top             =   8010
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   12150
      Picture         =   "frmListadoBotesReactivoEx.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7830
      Width           =   1410
   End
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Enabled         =   0   'False
      Height          =   960
      Left            =   4455
      Picture         =   "frmListadoBotesReactivoEx.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7830
      Width           =   1410
   End
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado de pedidos pendientes"
      Height          =   960
      Left            =   5895
      Picture         =   "frmListadoBotesReactivoEx.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7830
      Width           =   2880
   End
   Begin VB.CommandButton cmdTerminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Terminar Bote"
      Height          =   960
      Left            =   1541
      Picture         =   "frmListadoBotesReactivoEx.frx":1D68
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7830
      Width           =   1410
   End
   Begin VB.CommandButton cmdAbrir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Abrir Bote"
      Enabled         =   0   'False
      Height          =   960
      Left            =   90
      Picture         =   "frmListadoBotesReactivoEx.frx":2C32
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7830
      Width           =   1410
   End
   Begin VB.CommandButton cmdPedido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pedido"
      Height          =   960
      Left            =   2992
      Picture         =   "frmListadoBotesReactivoEx.frx":3AFC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7830
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
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
      Height          =   2115
      Left            =   45
      TabIndex        =   22
      Top             =   360
      Width           =   13485
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   8100
         TabIndex        =   35
         Top             =   1440
         Width           =   1335
         Begin VB.OptionButton opCaducado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   675
            TabIndex        =   37
            Top             =   180
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton opCaducado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   105
            TabIndex        =   36
            Top             =   180
            Width           =   555
         End
      End
      Begin MSDataListLib.DataCombo cmbre 
         Height          =   360
         Left            =   1770
         TabIndex        =   3
         Top             =   690
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   960
         TabIndex        =   32
         Top             =   1425
         Width           =   2235
         Begin VB.OptionButton opAbierto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   1575
            TabIndex        =   6
            Top             =   180
            Width           =   615
         End
         Begin VB.OptionButton opAbierto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   1035
            TabIndex        =   5
            Top             =   180
            Width           =   555
         End
         Begin VB.OptionButton opAbierto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.CheckBox chkTodosReactivos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10050
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   11835
         Picture         =   "frmListadoBotesReactivoEx.frx":43C6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   1410
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Height          =   600
         Left            =   10515
         TabIndex        =   23
         Top             =   1425
         Width           =   2880
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No anulados"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   12
            Top             =   210
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Anulados"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   1635
            TabIndex        =   13
            Top             =   210
            Width           =   1230
         End
      End
      Begin VB.CheckBox chkTodosBotes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10050
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbBotes 
         Height          =   360
         Left            =   1770
         TabIndex        =   1
         Top             =   300
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1770
         TabIndex        =   10
         Top             =   1080
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   50593793
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   4140
         TabIndex        =   11
         Top             =   1080
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   50593793
         CurrentDate     =   38002
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   4320
         TabIndex        =   33
         Top             =   1440
         Width           =   2325
         Begin VB.OptionButton opTerminado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   7
            Top             =   180
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton opTerminado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   1050
            TabIndex        =   8
            Top             =   180
            Width           =   555
         End
         Begin VB.OptionButton opTerminado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   1620
            TabIndex        =   9
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   6885
         TabIndex        =   38
         Top             =   1635
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   9675
         TabIndex        =   34
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Terminado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   3285
         TabIndex        =   31
         Top             =   1635
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abierto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   30
         Top             =   1635
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Reactivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3510
         TabIndex        =   26
         Top             =   1125
         Width           =   555
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recep. desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   25
         Top             =   1140
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código Bote"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   24
         Top             =   390
         Width           =   1185
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5010
      Left            =   90
      TabIndex        =   0
      Top             =   2790
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   8837
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Listado de Botes de Reactivos Externos"
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
      Index           =   4
      Left            =   45
      TabIndex        =   28
      Top             =   0
      Width           =   13485
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   27
      Top             =   2505
      Width           =   13485
   End
End
Attribute VB_Name = "frmListadoBotesReactivoEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub chkTodas_Click()
'    If chkTodas.Value = Checked Then
'        cmbBotesReactivoEx.Text = ""
'        cmbBotesReactivoEx.Enabled = False
'    Else
'        cmbBotesReactivoEx.Enabled = True
'    End If
'End Sub

'Private Sub chkTodos_Click()
'    If chkTodos.Value = Checked Then
'        cmbClientes.Text = ""
'        cmbClientes.Enabled = False
'    Else
'        cmbClientes.Enabled = True
'    End If
'End Sub


Private Sub chkTodosBotes_Click()
    If chkTodosBotes.Value = Checked Then
        cmbbotes.Text = ""
        cmbbotes.Enabled = False
    Else
        cmbbotes.Enabled = True
    End If
End Sub

Private Sub chkTodosReactivos_Click()
    If chkTodosReactivos.Value = Checked Then
        cmbre.Text = ""
        cmbre.Enabled = False
    Else
        cmbre.Enabled = True
    End If
End Sub

Private Sub cmdAbrir_Click()
    If lista.ListItems.Count > 0 Then
        Dim fecha As String
        fecha = InputBox("Introduzca la fecha de apertura para los botes marcados.", "Fecha apertura", Format(Date, "dd/mm/yyyy"))
        If fecha <> "" Then
            If IsDate(fecha) = False Then
                MsgBox "El formato de la fecha no es correcto.", vbCritical, App.Title
                Exit Sub
            End If
'        If MsgBox("¿Abrir los botes marcados con fecha de hoy?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim i As Integer
            Dim obe As New clsBotes_ex
            Dim se As Boolean
            se = False
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    se = True
                    obe.Abrir lista.ListItems(i).Text, fecha
                End If
            Next
            If se = True Then
                Call buscar
            Else
                MsgBox "No hay ningún bote marcado.", vbInformation, App.Title
            End If
        End If
    End If
End Sub

Private Sub cmdetiqueta_Click()
    frmPosicionPegatina.Show 1
    Dim fpegatina As String
    If pegatina <> 0 Then
        fpegatina = Format(pegatina, "00")
        On Error GoTo fallo
        Dim i As Integer
        ' Generamos los datos del listado
        Dim rs As New ADODB.Recordset
        rs.Fields.Append "c1", adChar, 5, adFldUpdatable
        rs.Open
        rs.AddNew
        rs("c1") = lista.ListItems(lista.SelectedItem.Index)
        rs.Update
        ' Generar Listado
        Dim Listado As New dataReactivo
        ' Ocultar controles
        For i = 1 To Listado.Sections("detalle").Controls.Count
            Listado.Sections("detalle").Controls(i).Visible = False
        Next
        ' Pegatina
        For i = 1 To Listado.Sections("detalle").Controls.Count
            If Left(Listado.Sections("detalle").Controls(i).Name, 3) = Trim("l" & fpegatina) Or _
               Left(Listado.Sections("detalle").Controls(i).Name, 3) = Trim("c" & fpegatina) Then
                    Listado.Sections("detalle").Controls(i).Visible = True
            End If
        Next
        If EMPRESA.getID_EMPRESA <> 3 Then  ' FQ
            Listado.Sections("detalle").Controls(Trim("logo" & pegatina)).Visible = True
            Set Listado.Sections("detalle").Controls(Trim("logo" & pegatina)).Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
        End If
        Dim obe As New clsBotes_ex
        Dim otb As New clsTipos_bote_ex
        Dim ore As New clsTipos_reactivo_ex
        obe.cargar (lista.ListItems(lista.SelectedItem.Index))
        otb.cargar (obe.getTIPO_BOTE_EX_ID)
        ore.cargar (otb.getTIPO_REACTIVO_EX_ID)
        With Listado.Sections("detalle")
'            If EMPRESA.getID_EMPRESA = 1 Then
'                .Controls(Trim("c" & fpegatina & "1")).Caption = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
'            Else
                .Controls(Trim("c" & fpegatina & "1")).Caption = lista.ListItems(lista.SelectedItem.Index)
'            End If
            .Controls(Trim("c" & fpegatina & "2")).Caption = ore.getNOMBRE
            .Controls(Trim("c" & fpegatina & "3")).Caption = Format(obe.getFECHA_RECEPCION, "dd/mm/yyyy")
            If obe.getFECHA_APERTURA = "" Then
                .Controls(Trim("c" & fpegatina & "4")).Caption = " "
            Else
                .Controls(Trim("c" & fpegatina & "4")).Caption = Format(obe.getFECHA_APERTURA, "dd/mm/yyyy")
            End If
            .Controls(Trim("c" & fpegatina & "5")).Caption = Format(obe.getFECHA_CADUCIDAD, "dd/mm/yyyy")
            .Controls(Trim("c" & fpegatina & "6")).Caption = obe.getLOTE
        End With
        If EMPRESA.getID_EMPRESA <> 3 Then  ' FQ
            If Dir(EMPLEADO.getFIRMA) <> "" Then
                Set Listado.Sections("detalle").Controls(Trim("c" & fpegatina & "7")).Picture = LoadPicture(EMPLEADO.getFIRMA)
            End If
        End If
        Set Listado.DataSource = rs
        Listado.Caption = "Pegatinas Reactivos Externos"
        Listado.WindowState = vbMaximized
        Listado.Show
        Set rs = Nothing
    End If
    Exit Sub
fallo:
    MsgBox "Error al generar el listado de documentos.", vbCritical, Err.Description
End Sub

Private Sub cmdListado_Click()
    frmREX_Pedidos_Listado.Show 1
End Sub

Private Sub cmdmanual_Click()
    If lista.ListItems.Count > 0 Then
        Dim consulta As String
        Dim nuevo As String
        nuevo = InputBox("Intro nuevo numero : ", App.Title)
        If nuevo <> "" Then
            consulta = "update botes_ex set id_bote_ex = " & nuevo & " where id_bote_ex = " & CLng(lista.ListItems(lista.SelectedItem.Index).Text)
            execute_bd consulta
            cmdBuscar_Click
        End If
    End If
End Sub

Private Sub cmdPedido_Click()
    frmREX_Bote_Pedido.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTerminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim fecha As String
        fecha = InputBox("Introduzca la fecha de cierre para los botes marcados.", "Fecha cierre", Format(Date, "dd/mm/yyyy"))
        If fecha <> "" Then
            If IsDate(fecha) = False Then
                MsgBox "El formato de la fecha no es correcto.", vbCritical, App.Title
                Exit Sub
            End If
'        If MsgBox("¿Terminar los botes marcados con fecha de hoy?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim i As Integer
            Dim obe As New clsBotes_ex
            Dim se As Boolean
            se = False
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    se = True
                    obe.Terminar lista.ListItems(i).Text, fecha
                End If
            Next
            If se = True Then
                Call buscar
            Else
                MsgBox "No hay ningún bote marcado.", vbInformation, App.Title
            End If
        End If
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 50
    Me.Top = 50
    cabecera
    cargar_combo cmbre, New clsTipos_reactivo_ex
    cargar_combo cmbbotes, New clsTipos_bote_ex
    fdesde = Date
    fhasta = Date
    If EMPRESA.getID_EMPRESA = 1 Then
        cmdmanual.Visible = True
    End If
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Número", 1000, lvwColumnLeft)
        .Tag = "Número"
    End With
    With lista.ColumnHeaders.Add(, , "Código", 1450, lvwColumnCenter)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Reactivo", 3500, lvwColumnLeft)
        .Tag = "Reactivo"
    End With
    With lista.ColumnHeaders.Add(, , "Recepción", 1100, lvwColumnCenter)
        .Tag = "Recepción"
    End With
    With lista.ColumnHeaders.Add(, , "Apertura", 1100, lvwColumnCenter)
        .Tag = "Apertura"
    End With
    With lista.ColumnHeaders.Add(, , "Terminado", 1100, lvwColumnCenter)
        .Tag = "Terminado"
    End With
    With lista.ColumnHeaders.Add(, , "Caducidad", 1100, lvwColumnCenter)
        .Tag = "Caducidad"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Lote", 1450, lvwColumnCenter)
        .Tag = "Lote"
    End With
    With lista.ColumnHeaders.Add(, , "Precio", 1250, lvwColumnRight)
        .Tag = "Precio"
    End With
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim consulta As String
    Dim strBote As String
    Dim strReactivo As String
    Dim strAbierto As String
'    Dim strCerrado As String
    Dim strAnulado As String
    Dim strCaducado As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim f_desde As String
    Dim f_hasta As String
    f_desde = Format(fdesde, "yyyy-mm-dd")
    f_hasta = Format(fhasta, "yyyy-mm-dd")
    ' Tipo de Bote
    strBote = ""
    If chkTodosBotes.Value = Unchecked Then
        If cmbbotes.Text = "" Then
            MsgBox "Debe seleccionar un codigo de Bote.", vbExclamation, App.Title
            Exit Sub
        End If
        strBote = " AND be.tipo_bote_ex_id=" & cmbbotes.BoundText
    End If
    ' Tipo reactivo
    strReactivo = ""
    If chkTodosReactivos.Value = Unchecked Then
        If cmbre.Text = "" Then
            MsgBox "Debe seleccionar un Reactivo.", vbExclamation, App.Title
            Exit Sub
        End If
        strReactivo = " AND tb.tipo_reactivo_ex_id = " & cmbre.BoundText
    End If
    ' Tipo
    strTipo = ""
    If opTipo(0).Value = True Then
        strTipo = " AND (mu.anulada is Null or mu.anulada = 0)"
    ElseIf opTipo(1).Value = True Then
        strTipo = " AND mu.anulada = 1"
    Else
        strTipo = " "
    End If
    ' Fechas
    Dim fecha_desde As String
    fecha_desde = " AND be.fecha_recepcion>='" & f_desde & "'"
    Dim fecha_hasta As String
    fecha_hasta = " AND be.fecha_recepcion<='" & f_hasta & "'"
    ' Abierto
    strAbierto = ""
    If opAbierto(1).Value = True Then
        strAbierto = " AND be.fecha_apertura <> 0"
    ElseIf opAbierto(2).Value = True Then
        strAbierto = " AND be.fecha_apertura = 0"
    End If
    ' Terminado
    strTerminado = ""
    If opTerminado(1).Value = True Then
        strTerminado = " AND be.fecha_fin <> 0"
    ElseIf opTerminado(2).Value = True Then
        strTerminado = " AND be.fecha_fin = 0"
    End If
    ' Caducado
    strCaducado = ""
    If opCaducado(0).Value = True Then ' Si caducados
        strCaducado = " AND be.fecha_caducidad < '" & Format(Date, "yyyy-mm-dd") & "'"
    ElseIf opCaducado(1).Value = True Then
        strCaducado = " AND be.fecha_caducidad > '" & Format(Date, "yyyy-mm-dd") & "'"
    End If
    ' Anulado
    strAnulado = ""
    If opTipo(0).Value = True Then
        strAnulado = " AND be.anulado = 0"
    Else
        strAnulado = " AND be.anulado = 1"
    End If
    ' Query
    consulta = "SELECT be.id_bote_ex, " & _
               "       tb.codigo, " & _
               "       tr.nombre, " & _
               "       be.fecha_recepcion, " & _
               "       be.fecha_apertura, " & _
               "       be.fecha_fin, " & _
               "       be.fecha_caducidad, " & _
               "       be.tipo_bote_ex_id, " & _
               "       be.LOTE, " & _
               "       tb.precio " & _
               " FROM BOTES_EX be, " & _
               "      TIPOS_BOTE_EX tb, " & _
               "      TIPOS_REACTIVO_EX tr " & _
               " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex " & _
               "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
                          strReactivo & _
                          strBote & _
                          fecha_desde & _
                          fecha_hasta & _
                          strAbierto & _
                          strTerminado & _
                          strCaducado & _
                          strAnulado & _
               " ORDER BY be.id_bote_ex desc"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , Format(rs.Fields(0), "00000"))
                .SubItems(1) = rs.Fields(1)
                If IsNull(rs.Fields(2)) Then
                    .SubItems(2) = ""
                Else
                    .SubItems(2) = rs.Fields(2)
                End If
                If IsNull(rs.Fields(3)) Then
                    .SubItems(3) = ""
                Else
                    .SubItems(3) = rs.Fields(3)
                End If
                If IsNull(rs.Fields(4)) Then
                    .SubItems(4) = ""
                Else
                    .SubItems(4) = rs.Fields(4)
                End If
                If IsNull(rs.Fields(5)) Then
                    .SubItems(5) = ""
                Else
                    .SubItems(5) = rs.Fields(5)
                End If
                If IsNull(rs.Fields(6)) Then
                    .SubItems(6) = ""
                Else
                    .SubItems(6) = rs.Fields(6)
                End If
                .SubItems(7) = rs(7)
                .SubItems(8) = rs(8)
                .SubItems(9) = Format(rs(9), "currency")
            End With
            rs.MoveNext
        Wend
        lblmsg.Caption = "Botes entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & " (Encontrados : " & rs.RecordCount & ")"
    Else
        lblmsg.Caption = "No existe ningun bote con esos criterios."
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Botes.", vbCritical, Err.Description
End Sub

Private Sub lista_Click()
    cmdetiqueta.Enabled = False
    cmdAbrir.Enabled = True
    cmdTerminar.Enabled = False
    If lista.ListItems.Count > 0 Then
        If opTipo(0).Value = True Then
            cmdetiqueta.Enabled = True
            If lista.ListItems(lista.SelectedItem.Index).SubItems(4) = "" Then
                cmdTerminar.Enabled = False
'                cmdAbrir.Enabled = True
            Else
                If lista.ListItems(lista.SelectedItem.Index).SubItems(5) = "" Then
                    cmdTerminar.Enabled = True
                End If
            End If
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

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gbotereactivoex = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
        frmREX_Bote.Show 1
        gbotereactivoex = 0
    End If
End Sub
