VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEquipoCuadernoAvisos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Mensajería"
   ClientHeight    =   7650
   ClientLeft      =   3270
   ClientTop       =   3060
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmEquipoCuadernoAvisos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   13361
      Caption         =   "Equipos con Mantenimiento/Calibración/Verificación"
      BackColor       =   16777215
      TextColor       =   255
      HeaderColor     =   8421504
      Object.Height          =   7575
      Begin VB.CheckBox chkfuera 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mostrar equipos Fuera de Servicio"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   840
         Value           =   1  'Checked
         Width           =   2805
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo"
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
         Height          =   1185
         Left            =   5160
         TabIndex        =   3
         Top             =   450
         Width           =   2205
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Verificaciones"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   7
            Top             =   720
            Width           =   1845
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Calibraciones"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   6
            Top             =   510
            Width           =   1815
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mantenimiento"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   5
            Top             =   930
            Width           =   1875
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   4
            Top             =   270
            Value           =   -1  'True
            Width           =   1725
         End
      End
      Begin VB.CheckBox chkSolo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mostrar todo lo pendiente"
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   570
         Width           =   2805
      End
      Begin MSComctlLib.ListView lista 
         Height          =   5385
         Left            =   45
         TabIndex        =   1
         Top             =   1695
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   9499
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "* Fuera de fecha"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   60
         TabIndex        =   11
         Top             =   7320
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "* Fuera de servicio"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   60
         TabIndex        =   10
         Top             =   7110
         Width           =   1575
      End
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5040
         TabIndex        =   9
         Top             =   7200
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmEquipoCuadernoAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ALTO_MIN = 500
Const ALTO_MAX = 7650
'Const ANCHO_MIN = 3990
'Const ANCHO_MAX = 10080

Private mvarblnSinAvisos As Boolean
Private Sub Carga()
    
    Dim rs As ADODB.RecordSet
    Dim mvarobjEquipos As New clsEquipos
    
   On Error GoTo Carga_Error
    
    Dim tipo As Integer
    If opTipo(0).value = True Then
        tipo = 0
    ElseIf opTipo(1).value = True Then
        tipo = 1
    ElseIf opTipo(2).value = True Then
        tipo = 2
    Else
        tipo = 3
    End If
    Set rs = mvarobjEquipos.ListadoCuadernoAvisosTotal(usuario.getID_EMPLEADO, chkSolo.value, tipo, chkfuera.value)
        
    lista.ListItems.Clear
'    On Error Resume Next
'    leido = True
    lbltotal = "Total : " & rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs("ID_EQUIPO"), "0000"))
              .SubItems(1) = rs("NOMBRE")
              .SubItems(2) = rs("RESPONSABLE_INT")
              .SubItems(3) = rs("FECHA_PREVISTA")
              .SubItems(4) = rs("TIPO")
              .SubItems(5) = rs("ID_TIPO")
              .SubItems(6) = rs("ID_EVENTO")
              If CInt(rs("FS")) = 1 Then ' Fuera de servicio
                colorear lista.ListItems.Count, vbBlue
              End If
              If rs("fecha_prevista") < Date Then
                colorear lista.ListItems.Count, vbRed
              End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    
    Else
        mvarblnSinAvisos = True
    End If
    
    Set mvarobjEquipos = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

Carga_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Carga of Formulario frmEquipoCuadernoAvisos"
End Sub

Private Sub chkfuera_Click()
    Carga
End Sub

Private Sub chkSolo_Click()
    Carga
End Sub

Private Sub ControlPanelXP1_Expand(State As Boolean)
    If State = False Then
        Me.Height = ALTO_MIN
    Else
        Me.Height = ALTO_MAX
    End If

End Sub
Private Sub Form_Load()
    log (Me.Name)
    mvarblnSinAvisos = False
    Me.Top = 1120
'    Me.Left = 3500
    Me.Left = 50
'    Me.Width = 9050
'    lista.Width = Me.ScaleWidth
    cabecera
    
    ' Primero Mira en los parámetros del sistema si el usuario está autorizado a verlo todo
    Dim oParam As New clsParametros
    oParam.Carga parametros.USUARIOS_CUADERNO_AVISOS_EQUIPOS, ""
    
    If InStr(1, Replace(oParam.getVALOR, " ", ""), "," & CStr(prmIdUsuario)) > 0 Or _
       InStr(1, Replace(oParam.getVALOR, " ", ""), CStr(prmIdUsuario) & ",") > 0 Then
        chkSolo.value = Checked
    End If
    
'    lista.Height = ALTO_MIN
    Carga
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEquipoCuadernoAvisos = Nothing
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nº Equipo", 1000, lvwColumnLeft
        .Add , , "Equipo", 3000, lvwColumnLeft
        .Add , , "Responsable", 0, lvwColumnLeft
        .Add , , "Fecha Prevista", 1200, lvwColumnLeft
        .Add , , "Tipo", 1500, lvwColumnLeft
        .Add , , "idTipo", 0, lvwColumnLeft
        .Add , , "idEvento", 0, lvwColumnLeft
    End With
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

'    If lista.Height = ALTO_MAX Then
'        lista.Height = ALTO_MIN
        'Me.Width = ANCHO_MIN
        'cmdtodos.Visible = False
'    Else
'        lista.Height = ALTO_MAX
        'cmdtodos.Visible = True
'    End If
End Sub

Private Sub cargar_mensaje()
    Dim strTipo As String
    Dim objfrm As Object
    Dim strIdEquipo As String
    Dim strFecha As String, strId_Evento As String
    Dim mvarobjEquipos As New clsEquipos
    
    strTipo = UCase(ClrStr(lista.ListItems(lista.SelectedItem.Index).SubItems(5), False, True))
    
    strIdEquipo = lista.ListItems(lista.SelectedItem.Index)
    strFecha = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
    strId_Evento = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
    
    Select Case strTipo
        Case "2"
            Set objfrm = New frmEquipoEdicionMtoFechasEdicion
        Case "0"
            Set objfrm = New frmEquipoEdicionCalibracion
            'strId_Evento = mvarobjEquipos.BuscarIdCalibracion(CDate(strFecha), strIdEquipo)
        Case "1"
            Set objfrm = New frmEquipoEdicionVerificacion
            'strId_Evento = mvarobjEquipos.BuscarIdVerificacion(CDate(strFecha), strIdEquipo)
    End Select
    
    objfrm.VieneDeCuaderno = True
    objfrm.IdEvento = CLng(strId_Evento)
    objfrm.FechaPrevista = CDate(strFecha)
    objfrm.idEquipo = CLng(strIdEquipo)
        
    objfrm.Show vbModal
    
    If objfrm.RESULTADO Then
        Carga
    End If
    
    Unload objfrm
    Set mvarobjEquipos = Nothing
    Set objfrm = Nothing
    
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cargar_mensaje
    End If
End Sub

Public Property Get SinAvisos() As Boolean

    SinAvisos = mvarblnSinAvisos

End Property

Public Property Let SinAvisos(ByVal blnSinAvisos As Boolean)

    mvarblnSinAvisos = blnSinAvisos

End Property

Private Sub opTipo_Click(Index As Integer)
    Carga
End Sub
Private Sub colorear(fila As Integer, color As Long)
    Dim i As Integer
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub

