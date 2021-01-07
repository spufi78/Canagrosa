VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSE_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Nuevo Sellante"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13170
   Icon            =   "frmSE_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13170
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Index           =   16
      Left            =   4575
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   8580
      Width           =   4605
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Index           =   15
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   8580
      Width           =   4500
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   14
      Left            =   7470
      TabIndex        =   18
      Top             =   7980
      Width           =   2000
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8400
      Width           =   1365
   End
   Begin MSDataListLib.DataCombo cmbunidades 
      Height          =   315
      Left            =   6255
      TabIndex        =   17
      Top             =   7995
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   11
      Left            =   4395
      TabIndex        =   15
      Top             =   7995
      Width           =   900
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   870
      Left            =   12030
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8400
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10935
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8400
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   12
      Left            =   5325
      TabIndex        =   16
      Top             =   7995
      Width           =   900
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   10
      Left            =   2385
      TabIndex        =   14
      Top             =   7995
      Width           =   2000
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   9
      Left            =   60
      TabIndex        =   13
      Top             =   7995
      Width           =   2310
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Sellante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4950
      Left            =   45
      TabIndex        =   25
      Top             =   360
      Width           =   13080
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Index           =   17
         Left            =   1530
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   4140
         Width           =   11475
      End
      Begin VB.CheckBox chkENAC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC"
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
         Left            =   11745
         TabIndex        =   41
         Top             =   3150
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Index           =   13
         Left            =   1530
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3480
         Width           =   11475
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   1530
         TabIndex        =   9
         Top             =   3150
         Width           =   10155
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   1530
         TabIndex        =   8
         Top             =   2835
         Width           =   11475
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1530
         TabIndex        =   7
         Top             =   2520
         Width           =   11475
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1530
         TabIndex        =   6
         Top             =   2205
         Width           =   11475
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1530
         TabIndex        =   5
         Top             =   1890
         Width           =   11475
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1530
         TabIndex        =   4
         Top             =   1575
         Width           =   11475
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1530
         TabIndex        =   3
         Top             =   1260
         Width           =   11475
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1530
         TabIndex        =   1
         Top             =   585
         Width           =   11475
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   0
         Top             =   270
         Width           =   11475
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   1530
         TabIndex        =   2
         Top             =   900
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones para datos específicos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   6
         Left            =   90
         TabIndex        =   42
         Top             =   4140
         Width           =   1410
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
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
         Index           =   5
         Left            =   90
         TabIndex        =   38
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
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
         Index           =   1
         Left            =   90
         TabIndex        =   36
         Top             =   945
         Width           =   600
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Preparación"
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
         Index           =   12
         Left            =   90
         TabIndex        =   34
         Top             =   2595
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Preparado por"
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
         Index           =   11
         Left            =   90
         TabIndex        =   33
         Top             =   1965
         Width           =   1215
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facility"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   32
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Preparation"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   31
         Top             =   2910
         Width           =   810
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
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
         Index           =   8
         Left            =   90
         TabIndex        =   30
         Top             =   3240
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Process"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   29
         Top             =   1635
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Test"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   630
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
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
         Index           =   2
         Left            =   90
         TabIndex        =   26
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
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
         Index           =   3
         Left            =   90
         TabIndex        =   28
         Top             =   1320
         Width           =   705
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2235
      Left            =   45
      TabIndex        =   12
      Top             =   5715
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   3942
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
   Begin pryCombo.miCombo cmbDeterminaciones 
      Height          =   330
      Left            =   9495
      TabIndex        =   19
      Top             =   7995
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   582
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Norma Criterio de Aceptación"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4620
      TabIndex        =   40
      Top             =   8340
      Width           =   3345
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterio de Aceptación"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   60
      TabIndex        =   39
      Top             =   8340
      Width           =   3345
   End
   Begin VB.Image cmdModificar 
      Height          =   435
      Left            =   12720
      Picture         =   "frmSE_Detalle.frx":08CA
      Stretch         =   -1  'True
      ToolTipText     =   "Modificar"
      Top             =   7140
      Width           =   450
   End
   Begin VB.Image imgbajar 
      Height          =   480
      Left            =   12720
      Picture         =   "frmSE_Detalle.frx":1194
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image imgsubir 
      Height          =   480
      Left            =   12720
      Picture         =   "frmSE_Detalle.frx":15D6
      Top             =   5775
      Width           =   480
   End
   Begin VB.Image cmddel1 
      Height          =   435
      Left            =   12720
      Picture         =   "frmSE_Detalle.frx":1A18
      Stretch         =   -1  'True
      ToolTipText     =   "Eliminar"
      Top             =   6720
      Width           =   450
   End
   Begin VB.Image cmdok1 
      Height          =   435
      Left            =   12720
      Picture         =   "frmSE_Detalle.frx":22E2
      Stretch         =   -1  'True
      ToolTipText     =   "Añadir"
      Top             =   7560
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Nuevo Sellante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   45
      TabIndex        =   24
      Top             =   45
      Width           =   13080
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listado de Ensayos a Realizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   45
      TabIndex        =   35
      Top             =   5370
      Width           =   13080
   End
End
Attribute VB_Name = "frmSE_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const columnas_lista As Integer = 11

Private Sub cmdHistorialCambios_Click()
    If gSE_Sellante <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_sellante
        frmHistorialCambios.PK_ID = gSE_Sellante
        frmHistorialCambios.PK_TITULO = "Sellante " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmddel1_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.selectedItem.Index
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If

    If validar_ensayo = True Then
       With lista.ListItems(lista.selectedItem.Index)
            .SubItems(1) = txtDatos(9)
            .SubItems(2) = txtDatos(10)
            .SubItems(3) = txtDatos(11)
            .SubItems(4) = txtDatos(12)
            .SubItems(5) = cmbunidades.Text
            If cmbunidades.Text = "" Then
                .SubItems(6) = 0
            Else
                .SubItems(6) = cmbunidades.BoundText
            End If
            .SubItems(7) = txtDatos(14)
            .SubItems(10) = txtDatos(15) ' Criterio
            .SubItems(11) = txtDatos(16) ' Norma del criterio
            .SubItems(8) = cmbDeterminaciones.getTEXTO
            If cmbDeterminaciones.getTEXTO = "" Then
                .SubItems(9) = 0
            Else
                .SubItems(9) = cmbDeterminaciones.getPK_SALIDA
            End If
       End With
       txtDatos(9) = ""
       txtDatos(10) = ""
       txtDatos(11) = ""
       txtDatos(12) = ""
       txtDatos(14) = ""
       txtDatos(15) = ""
       txtDatos(16) = ""
       
       cmbunidades.Text = ""
       cmbDeterminaciones.limpiar
       txtDatos(9).SetFocus
    End If
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      Dim oSellante As New clsSellantes
      With oSellante
            .setENSAYO = txtDatos(0)
            .setENSAYO_INGLES = txtDatos(1)
            .setCLIENTE_ID = cmbclientes.getPK_SALIDA
            .setPROCESO = txtDatos(2)
            .setPROCESO_INGLES = txtDatos(3)
            .setINSTALACION = txtDatos(4)
            .setINSTALACION_INGLES = txtDatos(5)
            .setPREPARACION = txtDatos(6)
            .setPREPARACION_INGLES = txtDatos(7)
            .setPRODUCTO = txtDatos(8)
            .setOBSERVACIONES = txtDatos(13)
            .setOBSERVACIONES_DE = txtDatos(17)
            .setENAC = chkENAC.Value
      End With
      Dim ohc As New clsHistorial_cambios
      If gSE_Sellante = 0 Then
        If MsgBox("Va a introducir el nuevo tipo de sellante. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            gSE_Sellante = oSellante.Insertar
            If gSE_Sellante > 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_sellante
                    .setIDENTIFICADOR = gSE_Sellante
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = HC_CREACION
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el tipo de sellante. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del sellante."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            If oSellante.Modificar(gSE_Sellante) = False Then
                Exit Sub
            Else
                With ohc
                    .setTIPO = HC_TIPOS.HC_sellante
                    .setIDENTIFICADOR = gSE_Sellante
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = Trim(MOTIVO)
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      Dim oSellante_ensayo As New clsSellantes_ensayos
     ' Eliminar ensayos
      oSellante_ensayo.Eliminar (gSE_Sellante)
      ' Insertar parámetros
      Dim i As Integer
      For i = 1 To lista.ListItems.Count
        With oSellante_ensayo
            .setSELLANTE_ID = gSE_Sellante
            .setORDEN = i
            .setENSAYO = lista.ListItems(i).SubItems(1)
            .setENSAYO_INGLES = lista.ListItems(i).SubItems(2)
            .setRANGO_INFERIOR = lista.ListItems(i).SubItems(3)
            .setRANGO_SUPERIOR = lista.ListItems(i).SubItems(4)
            .setUNIDAD_ID = lista.ListItems(i).SubItems(6)
            .setREFERENCIA = lista.ListItems(i).SubItems(7)
            .setCRITERIO = lista.ListItems(i).SubItems(10)
            .setNORMA_CRITERIO = lista.ListItems(i).SubItems(11)
            If lista.ListItems(i).SubItems(9) = "" Then
                .setTIPO_DETERMINACION_ID = 0
            Else
                .setTIPO_DETERMINACION_ID = lista.ListItems(i).SubItems(9)
            End If
            If lista.ListItems(i).Checked = True Then
                .setACTIVO = 1
            Else
                .setACTIVO = 0
            End If
            If .Insertar = 0 Then
                Exit Sub
            End If
        End With
      Next
      
      MsgBox "Actualizaciones realizadas correctamente.", vbOKOnly + vbInformation, App.Title
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave ("Error al insertar el sellante : " & Err.Description)
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    cargar_combos
    If gSE_Sellante <> 0 Then
        cargar_datos
    End If
End Sub

Private Sub cmdok1_Click()
    If validar_ensayo = True Then
       With lista.ListItems.Add(, , "0")
            .SubItems(1) = txtDatos(9)
            .SubItems(2) = txtDatos(10)
            .SubItems(3) = txtDatos(11)
            .SubItems(4) = txtDatos(12)
            .SubItems(5) = cmbunidades.Text
            If cmbunidades.Text = "" Then
                .SubItems(6) = 0
            Else
                .SubItems(6) = cmbunidades.BoundText
            End If
            .SubItems(7) = cmbDeterminaciones.getTEXTO
            If cmbDeterminaciones.getTEXTO = "" Then
                .SubItems(8) = 0
            Else
                .SubItems(8) = cmbDeterminaciones.getPK_SALIDA
            End If
            .SubItems(10) = txtDatos(15)
            .SubItems(11) = txtDatos(16)
            
       End With
       txtDatos(9) = ""
       txtDatos(10) = ""
       txtDatos(11) = ""
       txtDatos(12) = ""
       txtDatos(15) = ""
       txtDatos(16) = ""
       cmbunidades.Text = ""
       cmbDeterminaciones.limpiar
       txtDatos(9).SetFocus
    End If
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub imgbajar_Click()
   On Error GoTo imgbajar_Click_Error

    If lista.ListItems.Count > 0 Then
        If lista.selectedItem.Index < lista.ListItems.Count Then
            Dim aux As String
            Dim i As Integer
            For i = 1 To columnas_lista
                aux = lista.ListItems(lista.selectedItem.Index + 1).SubItems(i)
                lista.ListItems(lista.selectedItem.Index + 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
            Next
            Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
        End If
    End If

   On Error GoTo 0
   Exit Sub

imgbajar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgbajar_Click of Formulario frmCE_Ficha"
End Sub

Private Sub imgsubir_Click()
   On Error GoTo imgsubir_Click_Error

    If lista.ListItems.Count > 0 Then
        If lista.selectedItem.Index > 1 Then
            Dim aux As String
            Dim i As Integer
            For i = 1 To columnas_lista
                aux = lista.ListItems(lista.selectedItem.Index - 1).SubItems(i)
                lista.ListItems(lista.selectedItem.Index - 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
            Next
            Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index - 1)
        End If
    End If

   On Error GoTo 0
   Exit Sub

imgsubir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgsubir_Click of Formulario frmCE_Ficha"
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Orden", 280, lvwColumnLeft
        .Add , , "Ensayo", 2000, lvwColumnLeft
        .Add , , "Test", 2000, lvwColumnLeft
        .Add , , "R.Inferior", 900, lvwColumnCenter
        .Add , , "R.Superior", 900, lvwColumnCenter
        .Add , , "Unidades", 1200, lvwColumnCenter
        .Add , , "UNIDAD_ID", 1, lvwColumnCenter
        .Add , , "Norma Ensayo", 2000, lvwColumnCenter
        .Add , , "Tipo Determinación", 2500, lvwColumnCenter
        .Add , , "TIPO_DETERMINACION_ID", 1, lvwColumnCenter
        .Add , , "Criterio Aceptacion", 2000, lvwColumnCenter
        .Add , , "Norma C.A.", 2000, lvwColumnCenter
    End With
    
End Sub

Private Sub lista_Click()
   On Error GoTo lista_Click_Error

    If lista.ListItems.Count > 0 Then
        txtDatos(9) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtDatos(10) = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        txtDatos(11) = lista.ListItems(lista.selectedItem.Index).SubItems(3)
        txtDatos(12) = lista.ListItems(lista.selectedItem.Index).SubItems(4)
        txtDatos(14) = lista.ListItems(lista.selectedItem.Index).SubItems(7) ' Norma de ensayo
        txtDatos(15) = lista.ListItems(lista.selectedItem.Index).SubItems(10) ' Criterio
        txtDatos(16) = lista.ListItems(lista.selectedItem.Index).SubItems(11) ' Norma del criterio
        cmbunidades.BoundText = lista.ListItems(lista.selectedItem.Index).SubItems(6)
        If lista.ListItems(lista.selectedItem.Index).SubItems(9) <> "" Then
            If lista.ListItems(lista.selectedItem.Index).SubItems(9) <> 0 Then
                cmbDeterminaciones.MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(9)
            Else
                cmbDeterminaciones.limpiar
            End If
        Else
            cmbDeterminaciones.limpiar
        End If
    End If

   On Error GoTo 0
   Exit Sub

lista_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_Click of Formulario frmSE_Detalle"
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre al ensayo.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(8)) = "" Then
        MsgBox "Debe darle una descripción al producto.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbclientes.getPK_SALIDA = 0 Then
        MsgBox "Debe asignar un cliente al sellante.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function

Public Sub cargar_datos()
    Dim oSellante As New clsSellantes
    With oSellante
        If .Carga(gSE_Sellante) = True Then
            Label1(2).Caption = "Modificación Sellante : " & .getPRODUCTO
            Me.Caption = Label1(2).Caption
            txtDatos(0) = .getENSAYO
            txtDatos(1) = .getENSAYO_INGLES
            cmbclientes.MostrarElemento .getCLIENTE_ID
            txtDatos(2) = .getPROCESO
            txtDatos(3) = .getPROCESO_INGLES
            txtDatos(4) = .getINSTALACION
            txtDatos(5) = .getINSTALACION_INGLES
            txtDatos(6) = .getPREPARACION
            txtDatos(7) = .getPREPARACION_INGLES
            txtDatos(8) = .getPRODUCTO
            txtDatos(13) = .getOBSERVACIONES
            txtDatos(17) = .getOBSERVACIONES_DE
            chkENAC.Value = .getENAC
            ' Cargar Ensayos
            Dim oSellantes_ensayos As New clsSellantes_ensayos
            Dim rs As ADODB.Recordset
            Set rs = oSellantes_ensayos.Listado(gSE_Sellante)
            If rs.RecordCount > 0 Then
                Do
                    With lista.ListItems.Add(, , rs(0))
                         .SubItems(1) = rs(1)
                         .SubItems(2) = rs(2)
                         .SubItems(3) = rs(3)
                         .SubItems(4) = rs(4)
                         .SubItems(5) = rs(5)
                         .SubItems(6) = rs(6)
                         .SubItems(7) = rs(8) ' Referencia (NORMA ENSAYO)
                         .SubItems(10) = rs(9) ' Criterio de aceptacion
                         .SubItems(11) = rs(10) ' Norma del criterio
                         
                         If rs(7) = 0 Then
                            .SubItems(8) = ""
                            .SubItems(9) = 0
                         Else
                            Dim oTD As New clsTipos_determinacion
                            oTD.CargarTipoDeterminacion rs(7)
                            .SubItems(8) = oTD.getNOMBRE
                            .SubItems(9) = rs(7)
                            Set oTD = Nothing
                         End If
                         If rs(11) = 1 Then
                            .Checked = True
                         Else
                            .Checked = False
                         End If
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
        End If
    End With
    Set oSellante = Nothing
    Set oSellantes_ensayos = Nothing
End Sub

Public Function validar_ensayo() As Boolean
    validar_ensayo = True
    If Trim(txtDatos(9)) = "" Then
        MsgBox "Debe introducir una descripción del ensayo.", vbInformation, App.Title
        validar_ensayo = False
        Exit Function
    End If
End Function

Public Sub cargar_combos()
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    cargar_combo cmbunidades, New clsUnidades
    llenar_combo cmbDeterminaciones, New clsTipos_determinacion, 0, frmTD_Detalle, ""
End Sub
