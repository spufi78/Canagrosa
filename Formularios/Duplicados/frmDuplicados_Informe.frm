VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDuplicados_Informe 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Diferencia entre Duplicados"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   Icon            =   "frmDuplicados_Informe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   11340
   Begin VB.CommandButton cmdduplicados 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicados"
      Height          =   840
      Left            =   1260
      Picture         =   "frmDuplicados_Informe.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8055
      Width           =   1140
   End
   Begin VB.CommandButton cmdTD 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Determinación"
      Height          =   840
      Left            =   90
      Picture         =   "frmDuplicados_Informe.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Ir a la muestra anterior"
      Top             =   8055
      Width           =   1140
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8055
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   45
      TabIndex        =   5
      Top             =   855
      Width           =   11250
      Begin VB.CheckBox bRevisadas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo las que no están revisadas"
         Height          =   240
         Left            =   5760
         TabIndex        =   12
         Top             =   720
         Width           =   3930
      End
      Begin VB.CheckBox bFueraRango 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo las que tengan alguna determinacion fuera de rango"
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   270
         Width           =   3930
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   810
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1035
         TabIndex        =   0
         Top             =   405
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   60686337
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3510
         TabIndex        =   1
         Top             =   405
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   60686337
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Desde "
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   8
         Top             =   465
         Width           =   690
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   4
         Left            =   2970
         TabIndex        =   7
         Top             =   450
         Width           =   405
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5700
      Left            =   45
      TabIndex        =   3
      Top             =   2295
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   10054
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
      Caption         =   "Listado de tipos de determación que se han realizado por duplicado"
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
      Width           =   7095
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10755
      Picture         =   "frmDuplicados_Informe.frx":1A5E
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Se marcan en rojo los tipos de determinación cuya diferencia de duplicados esta fuera de rango"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   420
      Width           =   6750
   End
   Begin VB.Label lblMsg 
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
      TabIndex        =   6
      Top             =   1980
      Width           =   11280
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmDuplicados_Informe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdduplicados_Click()
    If lista.ListItems.Count > 0 Then
        frmDuplicados_Detalle.PK_ID_TIPO_DETERMINACION = lista.ListItems(lista.selectedItem.Index).Text
        frmDuplicados_Detalle.PK_DIF = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        frmDuplicados_Detalle.fecha_desde = fdesde
        frmDuplicados_Detalle.fecha_hasta = fhasta
        frmDuplicados_Detalle.Show 1
    End If
End Sub
Private Sub cmdTD_Click()
    If lista.ListItems.Count > 0 Then
        frmTD_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmTD_Detalle.Show 1
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    fdesde = Date - 7
    fhasta = Date
    cabecera
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 700, lvwColumnLeft
        .Add , , "Determinación", 6500, lvwColumnLeft
        .Add , , "% Dif. Máxima", 1200, lvwColumnCenter
        .Add , , "En Rango", 1200, lvwColumnCenter
        .Add , , "Fuera Rango", 1200, lvwColumnCenter
    End With
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim consulta As String
    lista.ListItems.Clear
    ' Tipo de muestra
    Dim oDeter As New clsDeterminaciones
    Dim rs As ADODB.Recordset
    Dim td_ant As Long
    Me.MousePointer = 11
    Dim objSI As ListSubItem
    Set rs = oDeter.lista_determinaciones_duplicadas(fdesde, fhasta, bFueraRango, bRevisadas)
    If rs.RecordCount > 0 Then
        lblMsg.Caption = "Cantidad de tipos de determinacion listadas: " & rs.RecordCount
        While Not rs.EOF
            If rs(3) = 1 Then
                b = True
            Else
                b = False
            End If
            With lista.ListItems.Add(, , rs.Fields(0)) ' id
                If b Then
                    .bold = True
                    .ForeColor = RGB(255, 0, 0)
                End If
                Set objSI = .ListSubItems.Add(, , rs(1)) ' DETER
                If b Then
                    objSI.bold = True
                    objSI.ForeColor = RGB(255, 0, 0)
                End If
                Set objSI = .ListSubItems.Add(, , rs(2)) ' DIF.DUPLICADO
                If b Then
                    objSI.bold = True
                    objSI.ForeColor = RGB(255, 0, 0)
                End If
                If rs(3) = 0 Then ' CANTIDAD EN RANGO
                    Set objSI = .ListSubItems.Add(, , rs(4))
                    If b Then
                        objSI.bold = True
                        objSI.ForeColor = RGB(255, 0, 0)
                    End If
                    Set objSI = .ListSubItems.Add(, , "")
                    If b Then
                        objSI.bold = True
                        objSI.ForeColor = RGB(255, 0, 0)
                    End If
                Else ' CANTIDAD FUERA RANGO
                    Set objSI = .ListSubItems.Add(, , "")
                    If b Then
                        objSI.bold = True
                        objSI.ForeColor = RGB(255, 0, 0)
                    End If
                    Set objSI = .ListSubItems.Add(, , rs(4))
                    If b Then
                        objSI.bold = True
                        objSI.ForeColor = RGB(255, 0, 0)
                    End If
                End If
            End With
            td_ant = rs(0)
            rs.MoveNext
            If Not rs.EOF Then
                If rs(0) = td_ant Then
                    If rs(3) = 0 Then
                        lista.ListItems(lista.ListItems.Count).SubItems(3) = rs(4)
                    Else
                        lista.ListItems(lista.ListItems.Count).SubItems(4) = rs(4)
                    End If
                    rs.MoveNext
                End If
            End If
        Wend
    Else
        lblMsg.Caption = "No existen registros con esos criterios."
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras.", vbCritical, Err.Description
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
        cmdduplicados_Click
    End If
End Sub
