VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmAlodine_Lote 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Gestión de nuevo Lote Suministrado de Alodine"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   Icon            =   "frmAlodine_Lote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmResultados 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   45
      TabIndex        =   30
      Top             =   4950
      Width           =   10230
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Registro de resultados"
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   0
         TabIndex        =   31
         Top             =   1770
         Width           =   10125
         Begin VB.TextBox txtParametros 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   30
            TabIndex        =   33
            Top             =   225
            Width           =   8025
         End
         Begin VB.TextBox txtParametros 
            Alignment       =   1  'Right Justify
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
            Height          =   330
            Index           =   1
            Left            =   8085
            TabIndex        =   32
            Top             =   225
            Width           =   1965
         End
      End
      Begin MSComctlLib.ListView parametros 
         Height          =   1455
         Left            =   0
         TabIndex        =   34
         Top             =   300
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   2566
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
      Begin VB.Label lbldeter 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Resultado de los Parámetros"
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
         Height          =   300
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   10215
      End
   End
   Begin VB.Frame frmBotones 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   45
      TabIndex        =   26
      Top             =   7470
      Width           =   10140
      Begin VB.CommandButton cmdEdiciones 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ediciones"
         Height          =   870
         Left            =   1125
         Picture         =   "frmAlodine_Lote.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   90
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdClientes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clientes"
         Enabled         =   0   'False
         Height          =   870
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   90
         Width           =   1050
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Height          =   870
         Left            =   7965
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   90
         Width           =   1050
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ESC-Salir"
         Height          =   870
         Left            =   9045
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   90
         Width           =   1050
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   540
      Left            =   60
      TabIndex        =   14
      Top             =   4350
      Width           =   9330
      Begin VB.TextBox txtanno 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5445
         TabIndex        =   16
         Top             =   150
         Width           =   465
      End
      Begin VB.TextBox txtmuestra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         TabIndex        =   15
         Top             =   150
         Width           =   1515
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   5911
         TabIndex        =   17
         Top             =   150
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   2004
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   225
         Index           =   0
         Left            =   4980
         TabIndex        =   19
         Top             =   210
         Width           =   435
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número general de la muestra a añadir "
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   18
         Top             =   195
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdEliminaMuestra 
      BackColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   9450
      Picture         =   "frmAlodine_Lote.frx":711C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2970
      Width           =   675
   End
   Begin VB.CommandButton cmdInsertaMuestra 
      BackColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   9450
      Picture         =   "frmAlodine_Lote.frx":79E6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4275
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   30
      TabIndex        =   4
      Top             =   405
      Width           =   10155
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   2
         Left            =   5085
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1395
         Width           =   1335
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   1
         Left            =   990
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   630
         Width           =   9075
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   1845
         TabIndex        =   1
         Top             =   1380
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker fecha_fabricacion 
         Height          =   330
         Left            =   1830
         TabIndex        =   2
         Top             =   1740
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_caducidad 
         Height          =   330
         Left            =   5070
         TabIndex        =   3
         Top             =   1725
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbAlodine 
         Height          =   315
         Left            =   990
         TabIndex        =   0
         Top             =   270
         Width           =   9060
         _ExtentX        =   15981
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
      Begin MSComCtl2.DTPicker fecha_impresion 
         Height          =   330
         Left            =   8355
         TabIndex        =   10
         Top             =   1725
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin XtremeSuiteControls.PushButton cmbNuevaEdicion 
         Height          =   300
         Left            =   6795
         TabIndex        =   37
         Top             =   1395
         Width           =   1590
         _Version        =   851970
         _ExtentX        =   2805
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Nueva Edición"
         Appearance      =   5
         Picture         =   "frmAlodine_Lote.frx":82B0
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición Certificado"
         Height          =   240
         Index           =   5
         Left            =   3375
         TabIndex        =   25
         Top             =   1425
         Width           =   1725
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   22
         Top             =   675
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de /Entrega"
         Height          =   240
         Index           =   1
         Left            =   6840
         TabIndex        =   11
         Top             =   1770
         Width           =   1500
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   9
         Top             =   330
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de Caducidad"
         Height          =   240
         Index           =   6
         Left            =   3360
         TabIndex        =   8
         Top             =   1785
         Width           =   1755
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de Fabricación"
         Height          =   240
         Index           =   4
         Left            =   75
         TabIndex        =   7
         Top             =   1800
         Width           =   1890
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Num.Lote Suministrado"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   5
         Top             =   1410
         Width           =   1725
      End
   End
   Begin MSComctlLib.ListView listaMuestras 
      Height          =   1365
      Left            =   60
      TabIndex        =   20
      Top             =   2970
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   2408
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Muestras realizadas"
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
      Height          =   255
      Left            =   45
      TabIndex        =   21
      Top             =   2670
      Width           =   10155
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Creación de nuevo Lote Suministrado de Alodine"
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
      Height          =   330
      Left            =   30
      TabIndex        =   6
      Top             =   45
      Width           =   10260
   End
End
Attribute VB_Name = "frmAlodine_Lote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbNuevaEdicion_Click()
   On Error GoTo cmbNuevaEdicion_Click_Error

    If MsgBox("¿Desea generar una nueva edición del certificado?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        frmMotivo.Show 1
        If Trim(MOTIVO) = "" Then
            MsgBox "Para generar una nueva edición es necesario introducir un motivo.", vbInformation, App.Title
            Exit Sub
        End If
        ' Aumentar edición
        Dim oalodine As New clsAlodine_lotes
        oalodine.aumentarEdicion glote
        Set oalodine = Nothing
        ' Historial de cambios (Edicion)
        Dim ohc As New clsHistorial_cambios
        With ohc
            .setTIPO = HC_TIPOS.HC_ALODINE_CERTIFICADOS
            .setIDENTIFICADOR = glote
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setIDENTIFICADOR_TEXTO = Me.Caption
            .setMOTIVO = Trim(MOTIVO)
            .Insertar
        End With
        imprimir glote, 20, False
        MsgBox "Nueva edición creada. La documentación esta siendo generada por el servidor. Espere unos segundos...", vbOKOnly + vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmbNuevaEdicion_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbNuevaEdicion_Click of Formulario frmAlodine_Lote"
End Sub

Private Sub cmdEdiciones_Click()
    If glote <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_ALODINE_CERTIFICADOS
        frmHistorialCambios.PK_ID = glote
        frmHistorialCambios.PK_TITULO = "Lote de Alodine " & Me.Caption
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmdClientes_Click()
    Dim LOTE As Long
    LOTE = glote
    gAlodine = cmbAlodine.BoundText
    frmAlodine_Clientes.Show 1
    glote = LOTE
End Sub
Private Sub cmbAlodine_Change()
    If glote = 0 Then
        CARGAR
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminaMuestra_Click()
    If listaMuestras.ListItems.Count > 0 Then
        listaMuestras.ListItems.Remove listaMuestras.selectedItem.Index
        txtMuestra = ""
    End If

End Sub

Private Sub cmdInsertaMuestra_Click()
    If listaMuestras.ListItems.Count = 1 Then
        MsgBox "Sólo puede añadirse 1 muestra realizada.", vbCritical, App.Title
        Exit Sub
    End If
    If txtMuestra <> "" Then
        If IsNumeric(txtMuestra) Then
            cargar_muestra_por_numero txtMuestra, txtanno
        End If
    End If

End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      ' Alodine
      Dim oAlodine_Lote As New clsAlodine_lotes
      Dim clsAlodine_resultados As New clsAlodine_resultados
      Dim LOTE As Long
      Dim i As Integer
      With oAlodine_Lote
           .setALODINE_ID = cmbAlodine.BoundText
           .setNUMERO_LOTE = txtDatos(0)
           .setFECHA_ALTA = Format(fecha_fabricacion, "yyyy-mm-dd")
           .setFECHA_CADUCIDAD = Format(fecha_caducidad, "yyyy-mm-dd")
           .setfecha_impresion = Format(fecha_impresion, "yyyy-mm-dd")
            ' Muestras
           Dim muestras As String
           muestras = "0"
           For i = 1 To listaMuestras.ListItems.Count
               muestras = muestras & listaMuestras.ListItems(i).SubItems(6)
           Next
'           If muestras <> "" Then
'             muestras = Left(muestras, Len(muestras) - 1)
'           End If
            .setMUESTRAS = muestras
           
           If glote = 0 Then
            LOTE = .Insertar
           Else
            .Modificar (glote)
            LOTE = glote
           End If
           If LOTE <> 0 Then
            'Parametros
            For i = 1 To parametros.ListItems.Count
              With clsAlodine_resultados
                  .setRESULTADO = parametros.ListItems(i).SubItems(1)
                  If glote = 0 Then
                    .setALODINE_ID = cmbAlodine.BoundText
                    .setLOTE_ID = LOTE
                    .setPARAMETRO_ID = parametros.ListItems(i).SubItems(2)
                    If .Insertar = 0 Then
                        Exit Sub
                    End If
                  Else
                    .Modificar glote, parametros.ListItems(i).SubItems(2)
                  End If
              End With
            Next
           Else
            Exit Sub
           End If
      End With
      If glote = 0 Then
          MsgBox "El lote se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
          glote = LOTE
          gAlodine = cmbAlodine.BoundText
          frmAlodine_Clientes.Show 1
      Else
          MsgBox "El lote se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave (Err.Description)
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_combos
    txtanno = Year(Date)
    cambiar.Max = Year(Date)
    If glote = 0 Then
        cmdClientes.Enabled = False
        fecha_fabricacion = Date
        fecha_caducidad = Date
        fecha_impresion = Date + 1
        cmbNuevaEdicion.visible = False
    Else
        cmdClientes.Enabled = True
        lbltitulo.BackColor = &H80C0FF
        lbltitulo = "Modificación de Lote Suministrado de Alodine"
        cmbAlodine.Enabled = False
'        txtdatos(0).Enabled = False
        cargar_lote
    End If
    ' Configurar nueva ventana sin parametros
    Dim op As New clsParametros
    op.Carga ALODINE_ULTIMO_LOTE_PARAMETROS, ""
    If glote = 0 Or glote > CInt(op.getVALOR) Then
        frmResultados.visible = False
        frmBotones.top = frmResultados.top
        Me.Height = Me.Height - frmResultados.Height
    End If
End Sub

Public Sub cabecera()
    ' Parametros
    With parametros.ColumnHeaders.Add(, , "Parámetro", 7755, lvwColumnLeft)
        .Tag = "Parámetro"
    End With
    With parametros.ColumnHeaders.Add(, , "Resultado", 2000, lvwColumnRight)
        .Tag = "Resultado"
    End With
    With parametros.ColumnHeaders.Add(, , "ID_PARAMETRO", 1, lvwColumnCenter)
        .Tag = "ID_PARAMETRO"
    End With
    ' Muestras
    With listaMuestras.ColumnHeaders
        .Add , , "Código", 900, lvwColumnLeft
        .Add , , "Cliente", 1800, lvwColumnLeft
        .Add , , "Tipo de Analisis/Solución", 2200, lvwColumnLeft
        .Add , , "Ref.Cliente", 2200, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "General", 800, lvwColumnCenter
        .Add , , "ID_MUESTRA", 1, lvwColumnCenter
    End With
    
End Sub

Private Sub listaMuestras_DblClick()
    If listaMuestras.ListItems.Count > 0 Then
        gmuestra = listaMuestras.ListItems(listaMuestras.selectedItem.Index).SubItems(6)
        frmVerMuestra.Show 1
    End If

End Sub

Private Sub parametros_Click()
    If parametros.ListItems.Count > 0 Then
        txtParametros(0) = parametros.ListItems(parametros.selectedItem.Index)
        txtParametros(1) = parametros.ListItems(parametros.selectedItem.Index).SubItems(1)
        txtParametros(1).SetFocus
    End If
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
Public Sub CARGAR()
    On Error GoTo fallo
    If cmbAlodine.BoundText <> "" Then
        ' Código de Alodine
        Dim oAlodine_Lote As New clsAlodine_lotes
        Dim oA As New clsAlodine
        oA.Carga cmbAlodine.BoundText
        With oAlodine_Lote
'            txtDatos(0) = .Proximo_NUMERO_LOTE(oA.getPRODUCTO, oA.getLOTE, Year(fecha_fabricacion))  ' & "/" & Year(fecha_fabricacion)
            txtDatos(0) = .Proximo_NUMERO_LOTE(oA.getID_ALODINE, Year(fecha_fabricacion))  ' & "/" & Year(fecha_fabricacion)
            txtDatos(1) = oA.getDESCRIPCION
        End With
        ' Caducidad
        Dim oalodine As New clsAlodine
        fecha_caducidad = fecha_fabricacion + oalodine.dias_caducidad(cmbAlodine.BoundText)
        ' Parámetros
        parametros.ListItems.Clear
        Dim oAlodine_Parametros As New clsAlodine_parametros
        Dim rs As ADODB.Recordset
        Set rs = oAlodine_Parametros.Listado_Parametros(cmbAlodine.BoundText)
        If rs.RecordCount <> 0 Then
           Do
             With parametros.ListItems.Add(, , rs(0))
                 .SubItems(1) = " "
                 .SubItems(2) = rs(5)
             End With
             rs.MoveNext
           Loop Until rs.EOF
        End If
        Set oAlodine_Parametros = Nothing
    End If
    Exit Sub
fallo:
    error_grave (Err.Description)
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbAlodine.BoundText = "" Then
        MsgBox "Debe seleccionar un tipo de alodine.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
'    Dim i As Integer
'    For i = 1 To parametros.ListItems.Count
'        If Trim(parametros.ListItems(i).SubItems(1)) = "" Then
'            MsgBox "Rellene los resultados de los parámetros.", vbExclamation, App.Title
'            validar = False
'            Exit Function
'        End If
'    Next
End Function

Public Sub cargar_combos()
    If glote = 0 Then
        cargar_combo cmbAlodine, New clsAlodine
    Else
        Dim rs As ADODB.Recordset
        Dim oA As New clsAlodine
        Set rs = oA.Listado_Combo_Todos
        Set cmbAlodine.RowSource = rs
        cmbAlodine.ListField = rs(1).Name
        cmbAlodine.BoundColumn = rs(0).Name
        Set rs = Nothing
    End If
End Sub

Private Sub txtParametros_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        If txtParametros(1) <> "" Then
            parametros.ListItems(parametros.selectedItem.Index).SubItems(1) = Trim(txtParametros(1))
            txtParametros(1) = ""
            ' Pasar al siguiente campo
            If parametros.ListItems.Count > parametros.selectedItem.Index Then
                Set parametros.selectedItem = parametros.ListItems(parametros.selectedItem.Index + 1)
                parametros_Click
            End If
        Else
            MsgBox "Introduzca un resultado.", vbCritical, App.Title
            txtParametros(1).SetFocus
        End If
    End If
End Sub
Private Sub txtparametros_LostFocus(Index As Integer)
    txtParametros(Index).BackColor = vbWhite
End Sub
Private Sub txtParametros_GotFocus(Index As Integer)
    txtParametros(Index).BackColor = &H80C0FF
    txtParametros(Index).SelStart = 0
    txtParametros(Index).SelLength = Len(txtParametros(Index))
End Sub

Public Sub cargar_lote()
    Dim oAlodine_Lote As New clsAlodine_lotes
    With oAlodine_Lote
        .Carga (glote)
        cmbAlodine.BoundText = .getALODINE_ID
'        txtDatos(0) = .getNUMERO_LOTE & "/" & Year(.getFECHA_ALTA)
        txtDatos(0) = .getNUMERO_LOTE
        fecha_fabricacion = .getFECHA_ALTA
        fecha_caducidad = .getFECHA_CADUCIDAD
        fecha_impresion = .getfecha_impresion
        txtDatos(2) = .getEDICION
        If .getEDICION > 1 Then
            cmdEdiciones.visible = True
        End If
        ' Parametros
        Dim oAlodine_resultados As New clsAlodine_resultados
        Dim rs As ADODB.Recordset
        Set rs = oAlodine_resultados.Resultados_Lote(glote)
        If rs.RecordCount <> 0 Then
            Do
                With parametros.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(3)
                    .SubItems(2) = rs(5)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        ' mUESTRAS
        If .getMUESTRAS <> "" Then
            cargar_muestras .getMUESTRAS
        End If
    End With
End Sub
Private Sub cargar_muestras(muestras As String)
    Dim consulta As String
   On Error GoTo cargar_muestras_Error

    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.tipo_analisis_id=ta.id_tipo_analisis AND " & _
                      " mu.id_muestra in (" & muestras & ")" & _
                      " order by mu.id_general desc"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Do
            With listaMuestras.ListItems.Add(, , rs(1))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(8)
                .SubItems(3) = rs.Fields(4)
                If Not IsNull(rs.Fields(5)) Then
                .SubItems(4) = rs.Fields(5)
                End If
                If Not IsNull(rs.Fields(9)) Then
                   .SubItems(5) = Format(rs.Fields(9), "00000")
                End If
                If Not IsNull(rs.Fields(6)) Then
                    .SubItems(6) = rs.Fields(6)
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Sub

cargar_muestras_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestras of Formulario frmAlodine_Alodine"

End Sub
Private Sub cargar_muestra_por_numero(muestra As Long, ANNO As Long)
    Dim consulta As String
   On Error GoTo cargar_muestra_por_numero_Error

    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                      " AND mu.id_general =" & muestra & _
                      " AND mu.anno = " & ANNO & _
                      " order by mu.id_general desc"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        ' Verificar si la muestra no existe en otro lote
        consulta = "select * from alodine_lotes where find_in_set(muestras," & rs(6) & ")"
        Dim rsAux As ADODB.Recordset
        Set rsAux = datos_bd(consulta)
        If rsAux.RecordCount > 0 Then
            MsgBox "ATENCION, LA MUESTRA INDICADA YA SE ENCUENTRA EN EL LOTE " & rsAux("NUMERO_LOTE") & "/" & Year(rsAux("FECHA_ALTA")), vbExclamation, App.Title
        End If
        Set rsAux = Nothing
        ' Fin Verificación
        Do
            With listaMuestras.ListItems.Add(, , rs(1))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(8)
                .SubItems(3) = rs.Fields(4)
                If Not IsNull(rs.Fields(5)) Then
                .SubItems(4) = rs.Fields(5)
                End If
                If Not IsNull(rs.Fields(9)) Then
                   .SubItems(5) = Format(rs.Fields(9), "00000")
                End If
                If Not IsNull(rs.Fields(6)) Then
                    .SubItems(6) = rs.Fields(6)
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Sub

cargar_muestra_por_numero_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestra_por_numero of Formulario frmAlodine_Alodine"

End Sub


