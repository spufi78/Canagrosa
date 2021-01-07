VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDetallePlantillaBano 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Plantilla de Baños"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmDetallePlantillaBano.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6750
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6750
      Width           =   1050
   End
   Begin VB.TextBox txtsel 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   6975
      Width           =   855
   End
   Begin VB.TextBox txtplantilla 
      Height          =   375
      Left            =   3690
      TabIndex        =   5
      Top             =   6735
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1365
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5535
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin MSDataListLib.DataCombo cmbBanos 
      Height          =   360
      Left            =   900
      TabIndex        =   3
      Top             =   6135
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccionados"
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
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   7035
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Baño"
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
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   6195
      Width           =   855
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   6570
   End
End
Attribute VB_Name = "frmDetallePlantillaBano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
   If cmbBanos.Text <> "" Then
    With lista.ListItems.Add(, , cmbBanos.BoundText)
         .SubItems(1) = cmbBanos.Text
    End With
    lista.ListItems(lista.ListItems.Count).Checked = True
   End If
   contar_sel
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
    Dim i As Integer
    Dim j As Integer
    If nueva_plantilla = 0 Then
        PLANTILLA = CInt(txtplantilla)
        ReDim plantilla_bano(CInt(txtsel))
        j = 1
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                plantilla_bano(j) = lista.ListItems(i)
                j = j + 1
            End If
        Next
        Unload Me
        Dim oform As New frmRecepcion
        oform.Show
        Set oform = Nothing
    Else
        Dim consulta As String
        If nueva_plantilla = 2 Then
            consulta = "delete from plantilla_banos where plantilla_id = " & PLANTILLA
            execute_bd consulta
        End If
        Dim opb As New clsPlantilla_banos
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                With opb
                    .setPLANTILLA_ID = PLANTILLA
                    .setBANO_ID = lista.ListItems(i)
                    .setORDEN_MUESTRA = i
                    If .Insertar = 0 Then
                        Exit Sub
                    End If
                End With
            End If
        Next
        consulta = "update plantillas_muestras set cantidad_muestras = " & CInt(txtsel) & " where id_plantilla = " & PLANTILLA
        execute_bd consulta
        MsgBox "Los baños de la plantilla se han almacenado correctamente.", vbInformation, App.Title
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    txtplantilla = PLANTILLA
    ReDim plantilla_bano(1)
    plantilla_bano(1) = 0
    cabecera
    cargar_banos
    cargar_lista
End Sub

Public Sub cargar_banos()
    Dim obanos As New clsBanos
    Dim oplantilla As New clsPlantillas_muestras
    oplantilla.CargaPlantilla (CInt(txtplantilla))
    lbltitulo = "Plantilla : " & oplantilla.getNOMBRE
    Set cmbBanos.RowSource = obanos.banos_cliente(oplantilla.getCLIENTE_ID, oplantilla.getTIPO_MUESTRA_ID)
    cmbBanos.ListField = "nombre"
    cmbBanos.BoundColumn = "id_bano"
    Set obanos = Nothing
    Set oplantilla = Nothing
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "NºBaño", 1100, lvwColumnLeft)
        .Tag = "NºBaño"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 5000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim opb As New clsPlantilla_banos
    lista.ListItems.Clear
    Dim oBANO As New clsBanos
    Set rs = opb.Listado_bano(CInt(txtplantilla))
    If rs.RecordCount <> 0 Then
        Do
'            If obano.cargar_bano(rs("bano_id")) = True Then
                With lista.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
                End With
                lista.ListItems(lista.ListItems.Count).Checked = True
'            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    contar_sel
    Set rs = Nothing
End Sub

Public Sub contar_sel()
    Dim i As Integer
    txtsel = "0"
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            txtsel = CInt(txtsel) + 1
        End If
    Next
End Sub
Private Sub lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    contar_sel
End Sub
