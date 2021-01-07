VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTDE_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Datos Específicos"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTDE_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   8280
   Begin VB.CheckBox chkob 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informar el parámetro obligatoriamente"
      Height          =   240
      Left            =   45
      TabIndex        =   9
      Top             =   6615
      Width           =   3120
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1132
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6930
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7155
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6930
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   855
      TabIndex        =   0
      Top             =   5850
      Width           =   7275
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4995
      Left            =   60
      TabIndex        =   1
      Top             =   765
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8811
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
   Begin MSDataListLib.DataCombo cmbUnidades 
      Height          =   315
      Left            =   855
      TabIndex        =   2
      Top             =   6210
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento de datos específicos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   420
      Width           =   2565
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   7695
      Picture         =   "frmTDE_Listado.frx":000C
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento de datos específicos"
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
      TabIndex        =   10
      Top             =   120
      Width           =   3750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unidad"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   6255
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   5895
      Width           =   555
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   8430
   End
End
Attribute VB_Name = "frmTDE_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
'    ElseIf cmbUnidades.BoundText = "" Then
'        MsgBox "La unidad no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a insertar el Dato Específico. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDE As New clsTipos_dato
            oDE.setNOMBRE = txtDatos(0)
            If cmbUnidades.BoundText = "" Then
                 oDE.setUNIDAD_ID = 0
            Else
                oDE.setUNIDAD_ID = cmbUnidades.BoundText
            End If
            If chkob.value = Checked Then
                oDE.setOBLIGATORIO = 1
            Else
                oDE.setOBLIGATORIO = 0
            End If
            oDE.Insertar
            chkob.value = Unchecked
            cargar_lista
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR el dato específico. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDE As New clsTipos_dato
            oDE.Eliminar (lista.ListItems(lista.selectedItem.Index).SubItems(3))
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar el tipo de unidad. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oTD As New clsTipos_dato
            oTD.setNOMBRE = txtDatos(0)
            If cmbUnidades.BoundText <> "" Then
                oTD.setUNIDAD_ID = cmbUnidades.BoundText
            End If
            If chkob.value = Checked Then
                oTD.setOBLIGATORIO = 1
            Else
                oTD.setOBLIGATORIO = 0
            End If
            oTD.Modificar (lista.ListItems(lista.selectedItem.Index).SubItems(3))
            cargar_lista
            txtDatos(0) = ""
            cmbUnidades.Text = ""
            chkob.value = Unchecked
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
    Me.Top = 200
    Me.Left = 200
    cargar_combo cmbUnidades, New clsUnidades
    cabecera
'    cargar_unidades
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oTipos_dato As New clsTipos_dato
    Dim ouni As New clsUnidades
    Set rs = oTipos_dato.Listado
    txtDatos(0) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           ouni.CARGAR (rs("unidad_id"))
           With lista.ListItems.Add(, , rs("nombre"))
            .SubItems(1) = ouni.getNOMBRE
            If rs("obligatorio") = 1 Then
                .SubItems(2) = "Si"
            Else
                .SubItems(2) = "No"
            End If
            .SubItems(3) = rs("id_tipo_dato")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oTipos_dato = Nothing
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
        cmbUnidades.Text = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        If lista.ListItems(lista.selectedItem.Index).SubItems(2) = "Si" Then
            chkob.value = Checked
        Else
            chkob.value = Unchecked
        End If
    End If
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Nombre", 4450, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Unidad", 1450, lvwColumnCenter)
        .Tag = "Unidad"
    End With
    With lista.ColumnHeaders.Add(, , "Obligatorio", 1450, lvwColumnCenter)
        .Tag = "Obligatorio"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 500, lvwColumnCenter)
        .Tag = "ID"
    End With
End Sub
