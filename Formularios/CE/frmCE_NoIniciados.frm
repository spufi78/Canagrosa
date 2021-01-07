VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCE_NoIniciados 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Mensajería"
   ClientHeight    =   6705
   ClientLeft      =   3270
   ClientTop       =   3060
   ClientWidth     =   9510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmCE_NoIniciados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   6585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9440
      _ExtentX        =   16642
      _ExtentY        =   11615
      Caption         =   "Ensayos de Eficacia No Iniciados"
      BackColor       =   16777215
      TextColor       =   255
      HeaderColor     =   8421504
      Object.Height          =   6585
      Begin VB.CheckBox chkSoloTiempo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mostrar sólo ensayos de eficacia"
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   510
         Width           =   2805
      End
      Begin MSComctlLib.ListView lista 
         Height          =   5595
         Left            =   45
         TabIndex        =   1
         Top             =   885
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   9869
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
      Begin VB.CheckBox chkfuera 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mostrar sólo los que están caducados o terminan hoy"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4980
         TabIndex        =   2
         Top             =   510
         Width           =   4275
      End
   End
End
Attribute VB_Name = "frmCE_NoIniciados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ALTO_MIN = 500
Const ALTO_MAX = 6705
'Const ANCHO_MIN = 3990
'Const ANCHO_MAX = 10080

Private Sub Carga()
   
    Dim rs As ADODB.RecordSet
    Dim oCE As New clsCe_recepcion
    
   On Error GoTo Carga_Error

    Set rs = oCE.Listado_NoIniciadas
        
    lista.ListItems.Clear
'    On Error Resume Next
'    leido = True
    Dim tiempo As Boolean
    Dim fuera As Boolean
    If rs.RecordCount > 0 Then
        Do
            tiempo = False
            fuera = False
            With lista.ListItems.Add(, , rs(0))
              .SubItems(1) = rs(1)
              .SubItems(2) = rs(2)
              .SubItems(3) = rs(3)
              .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
              ' Fecha Inicio rs(6)
              If Not IsNull(rs(6)) Then
                If rs(6) = "01-01-1900" Then
                    .SubItems(5) = "No Iniciado"
                Else
                    .SubItems(5) = Format(rs(6), "dd-mm-yyyy")
                End If
                tiempo = True
              End If
              ' Fecha fin rs(7)
              If Not IsNull(rs(7)) Then
                If rs(7) <> "01-01-1900" Then
                    .SubItems(6) = Format(rs(7), "dd-mm-yyyy")
                    If IsDate(rs(7)) Then
                       If CDate(Date) >= CDate(rs(7)) Then
                         colorear lista.ListItems.Count, vbRed
                         fuera = True
                       End If
                    End If
                End If
              End If
            End With
            ' Elimino si no es de tiempo y solo quiero los de tiempo
            If chkSoloTiempo = vbChecked And tiempo = False Then
                lista.ListItems.Remove lista.ListItems.Count
            Else
                ' Elimino si no esta fuera de plazo
                If chkfuera = vbChecked And fuera = False Then
                    lista.ListItems.Remove lista.ListItems.Count
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    Set oCE = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

Carga_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Carga of Formulario frmCE_NoIniciados"
End Sub

Private Sub chkfuera_Click()
    Carga
End Sub

Private Sub chkSoloTiempo_Click()
    If chkSoloTiempo.value = Checked Then
        chkfuera.Enabled = True
    Else
        chkfuera.Enabled = False
        chkfuera.value = Unchecked
    End If
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
    Me.Top = 1970
'    Me.Left = 3500
    Me.Left = 50
'    Me.Width = 9050
'    lista.Width = Me.ScaleWidth
    cabecera
'    lista.Height = ALTO_MIN
        Dim opar As New clsParametros
        opar.Carga parametros.DIAS_AVISO_CONTROLES_NOINICIADOS, ""
        If opar.getVALOR = "" Then
            dias = 7
        Else
            dias = opar.getVALOR
        End If
        ControlPanelXP1.Caption = "Muestras no cerradas y recepcionadas hace más de " & dias & " dias"
    Carga
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCE_NoIniciados = Nothing
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft
        .Add , , "NºGeneral", 1000, lvwColumnCenter
        .Add , , "Cliente", 2400, lvwColumnLeft
        .Add , , "Referencia", 2400, lvwColumnLeft
        .Add , , "F.Recepcion", 1000, lvwColumnCenter
        .Add , , "F.Comienzo", 1000, lvwColumnCenter
        .Add , , "F.Fin", 1000, lvwColumnCenter
    End With
End Sub
Private Sub cargar_mensaje()
    gmuestra = lista.ListItems(lista.SelectedItem.Index).Text
    frmVerMuestra.Show 1
'    With frmCE_Resultados
'        .PK_ID_MUESTRA = lista.ListItems(lista.SelectedItem.Index).Text
'        .Show 1
'    End With
    ControlPanelXP1.PanelOpen = False
    ControlPanelXP1.PanelOpen = True
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cargar_mensaje
    End If
End Sub

Private Sub colorear(fila As Integer, color As Long)
    Dim i As Integer
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub

