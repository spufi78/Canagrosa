VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmMuestras_Agrupar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agrupación de Muestras"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14010
   Icon            =   "frmMuestras_Agrupar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView listaNo 
      Height          =   6315
      Left            =   7560
      TabIndex        =   8
      Top             =   2655
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11139
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdbajar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignar"
      Height          =   780
      Left            =   6525
      Picture         =   "frmMuestras_Agrupar.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4860
      Width           =   1005
   End
   Begin VB.CommandButton cmdSubir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quitar"
      Height          =   780
      Left            =   6525
      Picture         =   "frmMuestras_Agrupar.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5670
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9000
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12915
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9000
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Muestra Agrupada"
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
      Height          =   1935
      Left            =   90
      TabIndex        =   5
      Top             =   360
      Width           =   13860
      Begin VB.TextBox txtReferencia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   10995
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   1
         Left            =   5205
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   315
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   0
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   315
         Width           =   960
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   315
         Width           =   1590
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1350
         TabIndex        =   15
         Top             =   720
         Width           =   12030
         _ExtentX        =   21220
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbTM 
         Height          =   330
         Left            =   1350
         TabIndex        =   16
         Top             =   1080
         Width           =   12030
         _ExtentX        =   21220
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fechaRecepcion 
         Height          =   330
         Left            =   8280
         TabIndex        =   21
         Top             =   315
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   51773441
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Recepción"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   6795
         TabIndex        =   22
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ref. muestra"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   135
         TabIndex        =   20
         Top             =   1470
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   18
         Top             =   750
         Width           =   1050
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de muestra"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   1125
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   3630
         TabIndex        =   14
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   360
         Width           =   825
      End
   End
   Begin MSComctlLib.ListView listaSi 
      Height          =   6315
      Left            =   90
      TabIndex        =   0
      Top             =   2655
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11139
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Muestras pendientes de agrupar"
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
      Left            =   7560
      TabIndex        =   9
      Top             =   2340
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agrupación de Muestras"
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
      Index           =   4
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   13875
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Muestras relacionadas con la muestra agrupada"
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
      Left            =   90
      TabIndex        =   6
      Top             =   2340
      Width           =   6390
   End
End
Attribute VB_Name = "frmMuestras_Agrupar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Public listaM As ListView
Private Enum COLS
    ID_MUESTRA = 0
    N_PARTICULAR = 1
    REFERENCIA_CLIENTE = 2
    N_GENERAL = 3
End Enum
Private Sub cmdbajar_Click()
    If listaNo.ListItems.Count > 0 Then
        With listaSi.ListItems.Add(, , listaNo.ListItems(listaNo.selectedItem.Index).Text)
            .SubItems(1) = listaNo.ListItems(listaNo.selectedItem.Index).SubItems(1)
            .SubItems(2) = listaNo.ListItems(listaNo.selectedItem.Index).SubItems(2)
            .SubItems(3) = listaNo.ListItems(listaNo.selectedItem.Index).SubItems(3)
        End With
        listaNo.ListItems.Remove listaNo.selectedItem.Index
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    If MsgBox("Va a realizar la agrupación de las muestras. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim i As Integer
        Dim oMuestra As New clsMuestra
        Dim lista As String
        On Error GoTo fallo
        Me.MousePointer = 11
        For i = 1 To listaSi.ListItems.Count
            If lista <> "" Then
                lista = lista & ","
            End If
            lista = lista & listaSi.ListItems(i).Text
        Next
        ' Informamos las muestras
        oMuestra.agruparMuestra PK, lista
        ' Iconos lista de muestras
        Dim j As Integer
        Dim objLitem As ListItem, objSitem As ListSubItem
        With listaM
            ' LISTASI
            If listaSi.ListItems.Count > 0 Then
                For j = 1 To listaSi.ListItems.Count
                    For i = 1 To .ListItems.Count
                      If .ListItems(i).SubItems(6) = listaSi.ListItems(j).Text Then
                        Set objLitem = .ListItems(i)
                        Set objSitem = objLitem.ListSubItems(10)
                        objSitem.ReportIcon = 11
                      End If
                    Next
                Next
            End If
            'LISTANO
            If listaNo.ListItems.Count > 0 Then
                For j = 1 To listaNo.ListItems.Count
                    For i = 1 To .ListItems.Count
                      If .ListItems(i).SubItems(6) = listaNo.ListItems(j).Text Then
                        Set objLitem = .ListItems(i)
                        Set objSitem = objLitem.ListSubItems(10)
                        objSitem.ReportIcon = vbNothing
                      End If
                    Next
                Next
            End If
            ' RAIZ
            For i = 1 To .ListItems.Count
              If .ListItems(i).SubItems(6) = PK Then
                Set objLitem = .ListItems(i)
                Set objSitem = objLitem.ListSubItems(10)
                If listaSi.ListItems.Count > 0 Then
                    objSitem.ReportIcon = 10
                Else
                    objSitem.ReportIcon = vbNothing
                End If
              End If
            Next
            
        End With
        MsgBox "Muestras agrupadas correctamente.", vbOKOnly + vbInformation, App.Title
        Set oMuestra = Nothing
        Me.MousePointer = 0
        Unload Me
    End If
    Exit Sub
fallo:
    Me.MousePointer = 0
    error_grave "Error al agrupar las muestras." & Err.Description
End Sub

Private Sub cmdSubir_Click()
    If listaSi.ListItems.Count > 0 Then
        With listaNo.ListItems.Add(, , listaSi.ListItems(listaSi.selectedItem.Index).Text)
            .SubItems(1) = listaSi.ListItems(listaSi.selectedItem.Index).SubItems(1)
            .SubItems(2) = listaSi.ListItems(listaSi.selectedItem.Index).SubItems(2)
            .SubItems(3) = listaSi.ListItems(listaSi.selectedItem.Index).SubItems(3)
        End With
        listaSi.ListItems.Remove listaSi.selectedItem.Index
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    cabecera
    cargar
End Sub
Private Sub cargar_combos()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbTM, New clsTipos_muestra, 0, frmTM_Detalle, ""
    cmbClientes.desactivar
    cmbTM.desactivar
End Sub
Private Sub cargar()
    Dim oMuestra As New clsMuestra
    Dim CODIGO As String
    With oMuestra
        .CargaMuestra PK
        txtNumero = .getID_GENERAL
        CODIGO = .CodigoParticular(.getID_MUESTRA)
        Pos = InStr(1, CODIGO, "-", vbTextCompare)
        txtCodigo(0) = Mid(CODIGO, 1, Pos - 1)
        fechaRecepcion = Format(.getFECHA_RECEPCION, "dd/mm/yyyy")
        txtCodigo(1) = .getID_PARTICULAR
        cmbClientes.MostrarElemento .getCLIENTE_ID
        cmbTM.MostrarElemento .getTIPO_MUESTRA_ID
        txtReferencia = .getREFERENCIA_CLIENTE
    End With
    Set oMuestra = Nothing
    cargarSi
    cargarNo
End Sub
Private Sub cargarSi()
    Dim oMuestra As New clsMuestra
   On Error GoTo cargar_muestras_Error

    Set rs = oMuestra.agrupadasListadoMuestra(PK)
    listaSi.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With listaSi.ListItems.Add(, , rs.Fields(0))
                .SubItems(1) = rs.Fields(1)
                .SubItems(2) = rs.Fields(2)
                .SubItems(3) = rs.Fields(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
   On Error GoTo 0
   Exit Sub
cargar_muestras_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargarSi of Formulario frmMuestras_Agrupar"
End Sub
Private Sub cargarNo()
    Dim oMuestra As New clsMuestra
   On Error GoTo cargar_muestras_Error

    Set rs = oMuestra.agrupadasListadoMuestraNoAsignadas(PK)
    listaNo.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With listaNo.ListItems.Add(, , rs.Fields(0))
                .SubItems(1) = rs.Fields(1)
                .SubItems(2) = rs.Fields(2)
                .SubItems(3) = rs.Fields(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
   On Error GoTo 0
   Exit Sub
cargar_muestras_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargarNo of Formulario frmMuestras_Agrupar"

End Sub
Private Sub cabecera()
    With listaSi.ColumnHeaders
        .Add , , "Id", 1, lvwColumnLeft
        .Add , , "NºEnsayo", 1000, lvwColumnCenter
        .Add , , "Ref.Cliente", 4100, lvwColumnLeft
        .Add , , "NºGeneral", 1000, lvwColumnCenter
    End With
    With listaNo.ColumnHeaders
        .Add , , "Id", 1, lvwColumnLeft
        .Add , , "NºEnsayo", 1000, lvwColumnCenter
        .Add , , "Ref.Cliente", 4100, lvwColumnLeft
        .Add , , "NºGeneral", 1000, lvwColumnCenter
    End With
End Sub
Private Sub listaNo_DblClick()
    cmdbajar_Click
End Sub
Private Sub listaSi_DblClick()
    cmdSubir_Click
End Sub
