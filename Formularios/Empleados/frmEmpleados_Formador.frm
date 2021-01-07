VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmEmpleados_Formador 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formación impartida por el empleado"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   Icon            =   "frmEmpleados_Formador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Left            =   45
      TabIndex        =   3
      Top             =   720
      Width           =   12570
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   810
         MaxLength       =   75
         TabIndex        =   6
         Top             =   360
         Width           =   2355
      End
      Begin VB.CommandButton cmdLimpiarCampos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar Campos"
         Height          =   780
         Left            =   11115
         Picture         =   "frmEmpleados_Formador.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1410
      End
      Begin pryCombo.miCombo cmbFormador 
         Height          =   330
         Left            =   6390
         TabIndex        =   8
         Top             =   360
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   582
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Persona Formada"
         Height          =   195
         Index           =   6
         Left            =   4815
         TabIndex        =   9
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "P.N.T."
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   450
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   885
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7785
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5970
      Left            =   75
      TabIndex        =   4
      Top             =   1740
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   10530
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Formación impartida por el empleado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   5970
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12105
      Picture         =   "frmEmpleados_Formador.frx":711C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Formación impartida por el empleado"
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
      TabIndex        =   0
      Top             =   45
      Width           =   3915
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   12825
   End
End
Attribute VB_Name = "frmEmpleados_Formador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub chkfiltro_Click()
    cargar_lista
End Sub

Private Sub cmbestados_Change()
    cargar_lista
End Sub
Private Sub chkmodalidad_Click(Index As Integer)
    cargar_lista

End Sub

Private Sub cmbFormador_change()
    cargar_lista
End Sub
Private Sub cmdLimpiarCampos_Click()
    txtdatos(0) = ""
    cmbFormador.Limpiar
    cargar_lista
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
        If lista.ListItems(lista.SelectedItem.Index).SubItems(4) = "" Then
            frmEmpleados_Cualificaciones_Nueva.EMPLEADO_ID = PK
        Else
            frmEmpleados_Cualificaciones_Nueva.EMPLEADO_ID = lista.ListItems(lista.SelectedItem.Index).Text
        End If
        frmEmpleados_Cualificaciones_Nueva.ID_CUALIFICACION = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        frmEmpleados_Cualificaciones_Nueva.Show 1
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    llenar_combo cmbFormador, New clsEmpleados, 0, frmEmpleados_Gestion, ""
    If PK > 0 Then
        cargar_lista
    End If
End Sub

Private Sub cargar_lista()
    Dim oempleado As New clsEmpleados
    If (oempleado.CARGAR(PK)) Then
        lbltitulo = "Formaciones del Empleado : " & oempleado.getNOMBRE
        Me.Caption = lbltitulo
        Dim oEF As New clsEmpleados_cualificaciones
        Dim rs As ADODB.RecordSet
        lista.ListItems.Clear
        If cmbFormador.getTEXTO = "" Then
            Set rs = oEF.Listado_Formacion(PK, txtdatos(0), 0)
        Else
            Set rs = oEF.Listado_Formacion(PK, txtdatos(0), cmbFormador.getPK_SALIDA)
        End If
        If rs.RecordCount > 0 Then
            Dim ant_pnt As Long
            Dim ant_pnt_nombre As String
            Do
                    If rs(6) = ant_pnt And ant_pnt_nombre = "" Then
                        If Not IsNull(rs(3)) Then
                            If Trim(rs(3)) <> "" And Format(rs(3), "yyyy-mm-dd") <> "1900-01-01" Then
                                lista.ListItems(lista.ListItems.Count).SubItems(3) = Format(rs(3), "dd-mm-yyyy")
                            End If
                        End If
                        If rs(5) = rs(0) Then
                            lista.ListItems(lista.ListItems.Count).SubItems(4) = "Puesta en Marcha"
                        Else
                            If Trim(rs(4)) = "" Then
                                lista.ListItems(lista.ListItems.Count).SubItems(4) = "No se ha formado a nadie."
                            Else
                                lista.ListItems(lista.ListItems.Count).SubItems(4) = rs(4)
                            End If
                        End If
                        
                    Else
                        With lista.ListItems.Add(, , rs(0))
                            .SubItems(1) = rs(1)
                            .SubItems(2) = rs(2)
                            If Not IsNull(rs(3)) Then
                                If Trim(rs(3)) <> "" And Format(rs(3), "yyyy-mm-dd") <> "1900-01-01" Then
                                    .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
                                End If
                            End If
                            If rs(5) = rs(0) Then
                                .SubItems(4) = "Puesta en Marcha"
                            Else
                                If Trim(rs(4)) = "" Then
                                    .SubItems(4) = "No se ha formado a nadie."
                                Else
                                    .SubItems(4) = rs(4)
                                End If
                            End If
                        End With
                    End If
                ant_pnt = rs(6)
                ant_pnt_nombre = rs(4)
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oEF = Nothing
        Set rs = Nothing
    End If
    Set oempleado = Nothing
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_EMPLEADO", 1, lvwColumnLeft
        .Add , , "ID_CUALIFICACION", 1, lvwColumnLeft
        .Add , , "P.N.T.", 6200, lvwColumnLeft
        .Add , , "Fecha", 1400, lvwColumnCenter
        .Add , , "P.Formada", 4500, lvwColumnCenter
    End With
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub
