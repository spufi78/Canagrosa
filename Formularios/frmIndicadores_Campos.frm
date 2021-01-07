VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmIndicadores_Campos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de campos de Indicadores"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "frmIndicadores_Campos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   6075
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4275
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4275
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   45
      TabIndex        =   11
      Top             =   3420
      Width           =   8205
      Begin MSDataListLib.DataCombo cmbValor 
         Height          =   360
         Left            =   1740
         TabIndex        =   8
         Top             =   270
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin VB.Label lblValor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   12
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   45
      TabIndex        =   3
      Top             =   360
      Width           =   8250
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
         Height          =   330
         Index           =   0
         Left            =   1740
         TabIndex        =   0
         Top             =   270
         Width           =   6345
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
         Height          =   1440
         Index           =   1
         Left            =   1740
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   660
         Width           =   6345
      End
      Begin MSDataListLib.DataCombo cmbFuncion 
         Height          =   360
         Left            =   1740
         TabIndex        =   2
         Top             =   2160
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Función"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   2190
         Width           =   855
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1170
         Width           =   1275
      End
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Valor de la Función"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   45
      TabIndex        =   13
      Top             =   3075
      Width           =   8205
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo Campo de Indicadores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   30
      TabIndex        =   7
      Top             =   15
      Width           =   8235
   End
End
Attribute VB_Name = "frmIndicadores_Campos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbFuncion_Change()
    If cmbFuncion.Text <> "" Then
        cmbValor.Text = ""
        Select Case cmbFuncion.BoundText
            Case 1, 5 ' T.M.
                cargar_combo cmbValor, New clsTipos_muestra
            Case 2, 6 ' Cliente
                cargar_combo cmbValor, New clsCliente
            Case 3, 4 ' Familias
                cargar_combo cmbValor, New clsFamilias
        End Select
    End If
End Sub


Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    If validar = True Then
        On Error GoTo fallo
        Dim oIndicadores_campos As New clsIndicadores_campos
        With oIndicadores_campos
            .setNOMBRE = txtDatos(0)
            .setDESCRIPCION = txtDatos(1)
            If cmbFuncion.Text <> "" Then
                .setFUNCION_ID = cmbFuncion.BoundText
            End If
            If cmbValor.Text <> "" Then
                .setDATOS = cmbValor.Text
            End If
            If gindicadores_campos = 0 Then
                If MsgBox("Va a introducir una nuevo campo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    .Insertar
                    MsgBox "El Campo se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
                End If
            Else
                If MsgBox("Va a modificar el campo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    .Modificar (gindicadores_campos)
                    MsgBox "El Campo se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
                End If
            End If
        End With
        Unload Me
    End If
    Exit Sub
fallo:
    MsgBox "Error al aceptar el campo.", vbCritical, App.Title
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo cmbFuncion, New clsIndicadores_funciones
    If gindicadores_campos <> 0 Then
        Cargar_Campo
    End If
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub Cargar_Campo()
    lbltitulo = "Modificación de Campo de Indicadores"
    Dim oIndicadores_campos As New clsIndicadores_campos
    With oIndicadores_campos
        If .Carga(gindicadores_campos) = True Then
            txtDatos(0) = .getNOMBRE
            txtDatos(1) = .getDESCRIPCION
            Dim oIndicadores_Funciones As New clsIndicadores_funciones
            If oIndicadores_Funciones.Carga(.getFUNCION_ID) = True Then
                cmbFuncion.Text = oIndicadores_Funciones.getNOMBRE
            End If
            ' DATOS
            Select Case .getFUNCION_ID
                Case 1 ' Por T.M.
                    cargar_combo cmbValor, New clsTipos_muestra
                Case 2 ' Clientes
                    cargar_combo cmbValor, New clsCliente
            End Select
            cmbValor.Text = .getDATOS
        End If
    End With
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre al campo.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function
