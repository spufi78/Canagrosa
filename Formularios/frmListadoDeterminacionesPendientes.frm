VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoDeterminacionesPendientes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Muestras con Determinaciones Pendientes"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   Icon            =   "frmListadoDeterminacionesPendientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   13320
   Begin VB.CommandButton cmdVerMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Muestra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   5085
      Picture         =   "frmListadoDeterminacionesPendientes.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7875
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12210
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7875
      Width           =   1050
   End
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   6165
      Picture         =   "frmListadoDeterminacionesPendientes.frx":1B3C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7875
      Width           =   1050
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
      Height          =   1110
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   13170
      Begin VB.TextBox txtanno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6300
         TabIndex        =   14
         Top             =   615
         Width           =   705
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
         Height          =   885
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   150
         Width           =   1155
      End
      Begin VB.TextBox txtp2 
         Appearance      =   0  'Flat
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
         Left            =   4680
         TabIndex        =   8
         Top             =   630
         Width           =   1065
      End
      Begin VB.TextBox txtp1 
         Appearance      =   0  'Flat
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
         Left            =   3240
         TabIndex        =   7
         Top             =   630
         Width           =   1065
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10140
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbMuestras 
         Height          =   360
         Left            =   1785
         TabIndex        =   5
         Top             =   195
         Width           =   8280
         _ExtentX        =   14605
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
      Begin MSComCtl2.UpDown cambiar 
         Height          =   375
         Left            =   7005
         TabIndex        =   15
         Top             =   615
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196613
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5850
         TabIndex        =   16
         Top             =   675
         Width           =   585
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
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
         Left            =   4410
         TabIndex        =   9
         Top             =   690
         Width           =   195
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de Ensayo Particular, desde"
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
         TabIndex        =   4
         Top             =   705
         Width           =   2955
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
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
         Left            =   90
         TabIndex        =   3
         Top             =   255
         Width           =   1545
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5925
      Left            =   45
      TabIndex        =   10
      Top             =   1905
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   10451
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
   Begin MSComctlLib.ListView muestras 
      Height          =   5925
      Left            =   5085
      TabIndex        =   19
      Top             =   1905
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10451
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   13230796
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
   Begin VB.Label lbltotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   60
      TabIndex        =   20
      Top             =   7920
      Width           =   4980
   End
   Begin VB.Label lblmsg2 
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
      Height          =   315
      Left            =   5085
      TabIndex        =   18
      Top             =   1560
      Width           =   8145
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Listado de Trabajo Pendiente"
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
      Height          =   330
      Index           =   4
      Left            =   60
      TabIndex        =   2
      Top             =   45
      Width           =   13155
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
      Height          =   315
      Left            =   45
      TabIndex        =   1
      Top             =   1560
      Width           =   5010
   End
End
Attribute VB_Name = "frmListadoDeterminacionesPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTodas_Click()
    txtp1 = ""
    txtp2 = ""
    If chkTodas.Value = Checked Then
        cmbMuestras.Text = ""
        cmbMuestras.Enabled = False
    Else
        cmbMuestras.Enabled = True
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdDeter_Click()
    If muestras.ListItems.Count > 0 Then
        gmuestra = muestras.ListItems(muestras.selectedItem.Index).SubItems(5)
        abrirRegistroMuestra gmuestra
'        frmDeterminaciones.Show 1
        gmuestra = 0
    End If
End Sub

Private Sub cmdVerMuestra_Click()
    If muestras.ListItems.Count > 0 Then
        gmuestra = muestras.ListItems(muestras.selectedItem.Index).SubItems(5)
        frmVerMuestra.Show 1
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    txtanno = Year(Date)
    cabecera
    cargar_combos
End Sub
Public Sub cabecera()
    ' Lista
    With lista.ColumnHeaders.Add(, , "Determinacion", 4000, lvwColumnLeft)
        .Tag = "Determinacion"
    End With
    With lista.ColumnHeaders.Add(, , "Total", 600, lvwColumnRight)
        .Tag = "Total"
    End With
    With lista.ColumnHeaders.Add(, , "TD", 1, lvwColumnRight)
        .Tag = "TD"
    End With
    ' Muestras
    With muestras.ColumnHeaders.Add(, , "Código", 800, lvwColumnLeft)
        .Tag = "Código"
    End With
    With muestras.ColumnHeaders.Add(, , "Recepcion", 1100, lvwColumnLeft)
        .Tag = "Recepcion"
    End With
    With muestras.ColumnHeaders.Add(, , "Cliente", 2200, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With muestras.ColumnHeaders.Add(, , "Ref.Cliente", 2800, lvwColumnLeft)
        .Tag = "Ref.Cliente"
    End With
    With muestras.ColumnHeaders.Add(, , "Número", 800, lvwColumnCenter)
        .Tag = "Número"
    End With
    With muestras.ColumnHeaders.Add(, , "Id", 1, lvwColumnLeft)
        .Tag = "Id"
    End With
End Sub
Public Sub cargar_combos()
    cargar_combo cmbMuestras, New clsTipos_muestra
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim consulta As String
    Dim strMuestra As String
    Dim strpar As String
    Dim strTipo As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    ' Tipo de muestra
    strMuestra = ""
    If chkTodas.Value = Unchecked Then
        If cmbMuestras.Text = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        strMuestra = " AND mu.tipo_muestra_id=" & cmbMuestras.BoundText
    End If
    ' Particular
    strpar = ""
    If txtp1 <> "" Or txtp2 <> "" Then
        If txtp1 = "" Or txtp2 = "" Then
            MsgBox "Debe completar los codigos de búsqueda.", vbInformation, App.Title
            Exit Sub
        Else
            If IsNumeric(txtp1) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp1.SetFocus
                Exit Sub
            End If
            If IsNumeric(txtp2) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp2.SetFocus
                Exit Sub
            End If
            strpar = " AND mu.id_particular between " & CLng(txtp1) & " and " & CLng(txtp2)
        End If
    End If
    'Tipo
    strTipo = ""
    consulta = "SELECT td.nombre, td.id_tipo_determinacion, count(*) " & _
               "  FROM muestras as mu, " & _
               "       determinaciones as d, " & _
               "       tipos_determinacion as td " & _
               " WHERE d.tipo_determinacion_id=td.id_tipo_determinacion AND " & _
               "       mu.id_muestra=d.muestra_id AND " & _
               "       ((d.resultado = '' or d.resultado IS NULL) " & _
               "     OR (d.resultado='--' and mu.CERRADA = 0  ) " & _
               "     or (ucase(d.resultado)='PENDIENTE')) " & _
               strMuestra & _
               strpar & _
               "   AND mu.anno = " & CInt(txtanno) & _
               "   AND mu.anulada = 0 " & _
               strTipo & _
               " group by d.tipo_determinacion_id"
'                              "     or (ucase(d.resultado)='PENDIENTE' and mu.CERRADA = 0 )) "

    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    Dim total As Long
    total = 0
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(2)
                .SubItems(2) = rs(1)
            End With
            total = total + rs(2)
            rs.MoveNext
        Wend
        lblMsg.Caption = "Total estadísticas."
        lista_Click
    Else
        lblMsg.Caption = "No existe nada pendiente con ese criterio."
    End If
    lbltotal = "Total Pendiente : " & total
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras." & Err.Description, vbCritical, Err.Description
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim i As Integer
    Dim consulta As String
    Dim strMuestra As String
    Dim strpar As String
    On Error GoTo fallo
    muestras.ListItems.Clear
    Dim rs As New ADODB.Recordset
    ' Tipo de muestra
    strMuestra = ""
    If chkTodas.Value = Unchecked Then
        If cmbMuestras.Text = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        strMuestra = " AND mu.tipo_muestra_id=" & cmbMuestras.BoundText
    End If
    ' Particular
    strpar = ""
    If txtp1 <> "" Or txtp2 <> "" Then
        If txtp1 = "" Or txtp2 = "" Then
            MsgBox "Debe completar los codigos de búsqueda.", vbInformation, App.Title
            Exit Sub
        Else
            If IsNumeric(txtp1) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp1.SetFocus
                Exit Sub
            End If
            If IsNumeric(txtp2) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp2.SetFocus
                Exit Sub
            End If
            strpar = " AND mu.id_particular between " & CLng(txtp1) & " and " & CLng(txtp2)
        End If
    End If
    consulta = "SELECT distinct concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "       mu.fecha_recepcion, " & _
               "       cl.nombre, " & _
               "       mu.referencia_cliente, " & _
               "       mu.id_general, " & _
               "       mu.id_muestra " & _
               "  FROM muestras as mu, " & _
               "       clientes as cl, " & _
               "       determinaciones as d, " & _
               "       tipos_determinacion as td, " & _
               "       tipos_muestra as tm " & _
               " WHERE d.tipo_determinacion_id=td.id_tipo_determinacion AND " & _
               "       mu.id_muestra=d.muestra_id AND " & _
               "       mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
               "       mu.cliente_id=cl.id_cliente AND " & _
               "       (d.resultado = '' or d.resultado IS NULL) " & _
               strMuestra & _
               strpar & _
               "   AND mu.anno = " & CInt(txtanno) & _
               "   AND mu.anulada = 0 " & _
               "   AND d.tipo_determinacion_id = " & lista.ListItems(lista.selectedItem.Index).SubItems(2) & _
               " order by mu.id_muestra"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        While Not rs.EOF
            With muestras.ListItems.Add(, , rs(0))
                .SubItems(1) = Format(rs(1), "dd-mm-yyyy")
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                .SubItems(4) = Format(rs(4), "00000")
                .SubItems(5) = rs(5)
            End With
            rs.MoveNext
        Wend
        lblmsg2.Caption = "Muestras con determinación pendiente."
    Else
        lblmsg2.Caption = "No existe nada pendiente con ese criterio."
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras." & Err.Description, vbCritical, Err.Description
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
Private Sub muestras_DblClick()
    cmdVerMuestra_Click
End Sub
