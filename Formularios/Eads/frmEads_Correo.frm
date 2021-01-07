VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEads_Correo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de histórico de Baños"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   Icon            =   "frmEads_Correo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   10500
   Begin VB.CommandButton cmdInforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe"
      Height          =   1005
      Left            =   2670
      Picture         =   "frmEads_Correo.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Previsualizar informe de ensayo"
      Top             =   5940
      Width           =   1275
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   9180
      Picture         =   "frmEads_Correo.frx":157C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5940
      Width           =   1275
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
      Height          =   1005
      Left            =   1357
      Picture         =   "frmEads_Correo.frx":1886
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5940
      Width           =   1275
   End
   Begin VB.CommandButton cmdVerMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Muestra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   45
      Picture         =   "frmEads_Correo.frx":2150
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5940
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterios de selección"
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
      Height          =   810
      Left            =   45
      TabIndex        =   0
      Top             =   375
      Width           =   10365
      Begin VB.TextBox txtanno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   315
         Width           =   795
      End
      Begin MSDataListLib.DataCombo cmbBanos 
         Height          =   360
         Left            =   855
         TabIndex        =   1
         Top             =   300
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   ""
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
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   8986
         TabIndex        =   7
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196613
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2099
         Min             =   1990
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   7650
         TabIndex        =   8
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
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
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   480
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4455
      Left            =   45
      TabIndex        =   3
      Top             =   1455
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      Caption         =   "Control de Histórico de Baños"
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
      Height          =   285
      Index           =   4
      Left            =   15
      TabIndex        =   5
      Top             =   45
      Width           =   10410
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Información de resultados sobre el baño seleccionado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   45
      TabIndex        =   4
      Top             =   1215
      Width           =   10380
   End
End
Attribute VB_Name = "frmEads_Correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBanos_Change()
    cargar_bano
End Sub

Private Sub cmdDeter_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.SelectedItem.Index).Text
        frmDeterminaciones.Show 1
    End If
End Sub

Private Sub cmdExcel_Click()
    On Error GoTo fallo
    Me.MousePointer = 11
    Dim XLA As Excel.Application
    Dim XLW As Excel.Workbook
    Dim XLS As Excel.Worksheet
    Set XLA = New Excel.Application
    Set XLW = XLA.Workbooks.Open(ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EADS\Historico " & cmbBanos.Text & "-" & Year(Date) & ".xls")
    Set XLS = XLW.Worksheets(1)
    ' Datos
'    XLS.Cells(fila + i, col) = 0
    XLW.Save
    XLA.Visible = True
    Set XLW = Nothing
    Set XLA = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se han producido errores al abrir la hoja excel: " & Err.Description, vbCritical, "FILA:" & fila & " COL:" & col
    XLA.Visible = True
    Set XLW = Nothing
    Set XLA = Nothing
End Sub

Private Sub cmdInforme_Click()
    gmuestra = lista.ListItems(lista.SelectedItem.Index).Text
    frmPrevisualizar.Show 1
End Sub

Private Sub cmdVerMuestra_Click()
    lista_DblClick
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 30
    Me.Top = 30
    txtanno = Year(Date)
    cabecera
    cargar_combo cmbBanos, New clsBanos_Control
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Código", 1200, lvwColumnCenter)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Parametros", 6000, lvwColumnLeft)
        .Tag = "Parametros"
    End With
End Sub

Public Sub cargar_bano()
    If cmbBanos.BoundText = "" Then
        Exit Sub
    End If
    ' Cargamos los parametros del baño
'    On Error GoTo fallo
    lista.ListItems.Clear
    Dim oRango_bano As New clsRangos_bano
    Dim rs As ADODB.Recordset
    Set rs = oRango_bano.Listado_determinaciones_bano(cmbBanos.BoundText)
    Dim i As Integer
    For i = lista.ColumnHeaders.Count To 4 Step -1
        lista.ColumnHeaders.Remove i
    Next
    i = 3
    Dim pos(3 To 30) As Integer
    If rs.RecordCount <> 0 Then
        Do
            With lista.ColumnHeaders.Add(, , rs(0), 1800, lvwColumnRight)
                 .Tag = rs(0)
            End With
            pos(i) = rs(1)
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
        ' Resultados
        Dim obano As New clsBanos
        On Error Resume Next
        Set rs = obano.Resultados_Banos(cmbBanos.BoundText, txtanno)
        Dim muestra As String
        If rs.RecordCount <> 0 Then
            muestra = ""
            Do
                With lista.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                    .SubItems(2) = rs(2)
                    muestra = rs(0)
                    Do
                        For j = 3 To 30
                           If rs(4) = pos(j) Then ' ID_DETER
                                If Not rs(3) <> "" And Not IsNull(rs(3)) Then ' resultado
                                   .SubItems(j) = " "
                                Else
                                   .SubItems(j) = Replace(rs(3), ".", ",")
                                End If
                                Exit For
                           End If
                        Next
                       If rs.EOF = False Then
                           rs.MoveNext
                       End If
                    Loop Until muestra <> rs(0)
                End With
            Loop Until rs.EOF
        End If
    End If
    Exit Sub
fallo:
    MsgBox "Error al recuperar los resultados de los baños.", vbCritical, App.Title
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.SelectedItem.Index).Text
        frmVerMuestra.Show 1
    End If

End Sub

Private Sub UpDown1_Change()
    cargar_bano
End Sub
