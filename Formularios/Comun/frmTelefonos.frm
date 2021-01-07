VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmTelefonos 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Cambio Usuario"
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin Geslab.ControlPanelXP cpReactivos 
      Height          =   6690
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   11800
      Caption         =   "Extensiones Telefónicas"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   6690
      Begin MSComctlLib.ListView lista 
         Height          =   6105
         Left            =   60
         TabIndex        =   1
         Top             =   450
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   10769
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin Geslab.ControlPanelXP panelModificaciones 
      Height          =   6450
      Left            =   45
      TabIndex        =   2
      Top             =   405
      Visible         =   0   'False
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   11377
      Caption         =   "Listado de últimas modificaciones realizadas."
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   6450
      Begin RichTextLib.RichTextBox texto 
         Height          =   5955
         Left            =   45
         TabIndex        =   3
         Top             =   405
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   10504
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmTelefonos.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Geslab.ControlPanelXP panelFirmas 
      Height          =   7035
      Left            =   45
      TabIndex        =   4
      Top             =   405
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   12409
      Caption         =   "Firmas pendientes"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   7035
      Begin VB.CheckBox chkFirmasTodas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar las de todos los usuarios"
         Height          =   240
         Left            =   45
         TabIndex        =   10
         Top             =   6705
         Width           =   2760
      End
      Begin VB.Timer Timer1 
         Interval        =   30000
         Left            =   6435
         Top             =   6075
      End
      Begin MSComctlLib.ListView listaFirmas 
         Height          =   6240
         Left            =   45
         TabIndex        =   5
         Top             =   450
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   11007
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
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
   End
   Begin Geslab.ControlPanelXP panelAcciones 
      Height          =   6630
      Left            =   45
      TabIndex        =   6
      Top             =   835
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   11695
      Caption         =   "Acciones Correctivas / Preventivas"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   6630
      Begin XtremeSuiteControls.PushButton cmdVerProcn 
         Height          =   300
         Left            =   2610
         TabIndex        =   9
         Top             =   6255
         Width           =   2355
         _Version        =   851970
         _ExtentX        =   4154
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Ver Proc. N.C."
         BackColor       =   12632256
         Appearance      =   5
         Picture         =   "frmTelefonos.frx":007C
      End
      Begin XtremeSuiteControls.PushButton cmdVerAccion 
         Height          =   300
         Left            =   45
         TabIndex        =   8
         Top             =   6255
         Width           =   2535
         _Version        =   851970
         _ExtentX        =   4471
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Ver Acción"
         BackColor       =   12632256
         Appearance      =   5
         Picture         =   "frmTelefonos.frx":68DE
      End
      Begin VB.Timer TimerAcciones 
         Interval        =   30000
         Left            =   6390
         Top             =   6075
      End
      Begin MSComctlLib.ListView listaAcciones 
         Height          =   5745
         Left            =   45
         TabIndex        =   7
         Top             =   450
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   10134
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
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
   End
End
Attribute VB_Name = "frmTelefonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkFirmasTodas_Click()
    cargar_lista_firmas
End Sub

Private Sub cmdVerAccion_Click()
    If listaAcciones.ListItems.Count = 0 Then Exit Sub
    
    Dim objfrm As New frmProcNCEdicion_AccionCorrectiva
        
    objfrm.PK = listaAcciones.ListItems(listaAcciones.selectedItem.Index).Text
    objfrm.PK_PNC = listaAcciones.ListItems(listaAcciones.selectedItem.Index).SubItems(1)
    objfrm.NivelAcceso = ACCESO_TOTAL
    objfrm.estado_pnc = listaAcciones.ListItems(listaAcciones.selectedItem.Index).SubItems(5)
    
    objfrm.Show vbModal
    
    Unload objfrm
    Set objfrm = Nothing
    
    cargar_lista_acciones

End Sub

Private Sub cmdVerProcn_Click()
    If listaAcciones.ListItems.Count = 0 Then Exit Sub
    Dim objfrm As frmProcNCEdicion
    Set objfrm = New frmProcNCEdicion
    objfrm.PK = listaAcciones.ListItems(listaAcciones.selectedItem.Index).SubItems(1)
    objfrm.Show vbModal
    Unload objfrm
    Set objfrm = Nothing
End Sub

Private Sub cpReactivos_Expand(State As Boolean)
    comprobarAlto
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = Screen.Width - Me.Width - frmMenu.ButtonBar.Width - 500
    Me.top = 200
    cabecera
    cargar_lista
'M0996-I
    cargar_lista_firmas
'M0996-F
    cargar_lista_acciones
    comprobarAlto
    On Error Resume Next
    texto.LoadFile ReadINI(App.Path + "\config.ini", "Documentos", "Cambios"), 0
    If USUARIO.getPER_RFI = 0 Then
        chkFirmasTodas.Enabled = False
    End If
End Sub

'MANTIS-820-I
'M1127-I Recuperar lo que estaba

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oDeco As New clsDecodificadora
    Set rs = oDeco.Listado_por_Codigo(DECODIFICADORA.DECO_EXTENSIONES_TELEFONO)
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("DESCRIPCION"))
            .SubItems(1) = rs("PARAMETROS")
            .SubItems(2) = rs("VALOR")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
Private Sub cargar_lista_acciones()
    Dim rs As ADODB.Recordset
    Dim oAccCorrectoras As New clsProcNcAccionCorrectora
   On Error GoTo cargar_lista_acciones_Error

    listaAcciones.ListItems.Clear
    Set rs = oAccCorrectoras.ListadoAccionesPendientes()
    If rs.RecordCount <> 0 Then
        Do
'           panelAcciones.PanelOpen = True
           panelAcciones.Caption = "Tiene acciones correctivas / preventivas pendientes (" & rs.RecordCount & ")"
           panelAcciones.HeaderColor = vbRed
           comprobarAlto
           
           With listaAcciones.ListItems.Add(, , rs("ID_ACCION_CORRECTIVA"))
            .SubItems(1) = Format(rs("ID_PROCNC"), "000000")
            .SubItems(2) = rs("TITULO")
            .SubItems(3) = rs("RESPONSABLE")
            .SubItems(4) = Format(rs("FECHA_PREVISTA"), "dd/mm/yyyy")
            .SubItems(5) = rs("ESTADO_ID")
           End With
           rs.MoveNext
        Loop Until rs.EOF
        Dim Col As Integer
        Dim fila As Integer
        For fila = 1 To listaAcciones.ListItems.Count
            If Format(listaAcciones.ListItems(fila).SubItems(4), "yyyy-mm-dd") <= Format(Date, "yyyy-mm-dd") Then
                listaAcciones.ListItems(fila).ForeColor = vbRed
                For Col = 1 To listaAcciones.ColumnHeaders.Count - 1
                    listaAcciones.ListItems(fila).ListSubItems(Col).ForeColor = vbRed
                Next
            End If
        Next
    Else
        panelAcciones.Caption = "No tiene acciones correctivas / preventivas pendientes"
        panelAcciones.PanelOpen = False
        panelAcciones.HeaderColor = &H808080
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cargar_lista_acciones_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista_acciones of Formulario frmTelefonos"
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
'MANTIS-820-I
        '.Add , , "Ext.", 1200, lvwColumnLeft
        '.Add , , "Propietario.", 2320, lvwColumnLeft
        .Add , , "Departamento/Usuario", 3250, lvwColumnLeft
        .Add , , "Teléfono", 1500, lvwColumnLeft
        .Add , , "Ext.", 1200, lvwColumnLeft
'MANTIS-820-F
'M0996-I
    End With
    
    With listaFirmas.ColumnHeaders
        .Add , , "ID_Firma", 0, lvwColumnLeft
        .Add , , "Empleado", 1200, lvwColumnLeft
        .Add , , "Tipo", 2200, lvwColumnLeft
        .Add , , "Descripción", 3200, lvwColumnLeft
    End With
'M0996-F
    With listaAcciones.ColumnHeaders
        .Add , , "id_accion", 0, lvwColumnLeft
        .Add , , "NºP.N.C.", 750, lvwColumnCenter
        .Add , , "Titulo", 3000, lvwColumnLeft
        .Add , , "Responsable", 1750, lvwColumnLeft
        .Add , , "F.Prevista", 1050, lvwColumnLeft
        .Add , , "ESTADO_ID", 1, lvwColumnLeft
    End With
End Sub

'M0996-I
Public Sub cargar_lista_firmas()
    Dim rs As ADODB.Recordset
    
    Dim oempleado As New clsEmpleados
'    Dim ofirmante As New clsEmpleados
    Dim oFirmas As New clsFirmas
    
    listaFirmas.ListItems.Clear
    
    Dim TOBJETO As Integer
    
    If oempleado.CARGAR_POR_USUARIO(USUARIO.getID_EMPLEADO) = True Then
    
    
    Set rs = oFirmas.ListadoUsuario(oempleado.getID_EMPLEADO, chkFirmasTodas.Value)
    
    If rs.RecordCount <> 0 Then
'        panelFirmas.PanelOpen = True
        panelFirmas.Caption = "Tiene firmas pendientes (" & rs.RecordCount & ")"
        panelFirmas.HeaderColor = vbRed
        Do
           With listaFirmas.ListItems.Add(, , rs("ID_FIRMA"))
                If Not IsNull(rs(1)) Then
                   .SubItems(1) = rs(1)
                Else
                   .SubItems(1) = ""
                End If
                If Not IsNull(rs(2)) Then
                    .SubItems(2) = rs(2)
                End If
                If Not IsNull(rs(3)) Then
                    .SubItems(3) = rs(3)
                End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    Else
        panelFirmas.Caption = "No tiene firmas pendientes"
        panelFirmas.PanelOpen = False
        panelFirmas.HeaderColor = &H808080
    End If
    End If
    Set rs = Nothing
   ' Set oEmpleado = Nothing
    
End Sub

Private Sub listaAcciones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If listaAcciones.ListItems.Count > 0 Then
     listaAcciones.SortKey = ColumnHeader.Index - 1
     If listaAcciones.SortOrder = 0 Then
        listaAcciones.SortOrder = 1
     Else
        listaAcciones.SortOrder = 0
     End If
     listaAcciones.Sorted = True
   End If

End Sub

Private Sub listaAcciones_DblClick()
    cmdVerAccion_Click
End Sub
Private Sub colorear(lista As ListView, fila As Integer, color As Long)
    Dim i As Integer
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
    lista.Refresh
End Sub
Private Sub listaFirmas_DblClick()

    If listaFirmas.ListItems.Count > 0 Then
    
        frmFirmas.ID_FIRMA = listaFirmas.ListItems(listaFirmas.selectedItem.Index).Text
        frmFirmas.Show
'        panelFirmas.PanelOpen = False
        
    End If

End Sub
'M0996-F

Private Sub comprobarAlto()
    If panelModificaciones.PanelOpen = False And panelFirmas.PanelOpen = False And cpReactivos.PanelOpen = False And panelAcciones.PanelOpen = False Then
        Me.Height = 2040
    Else
        Me.Height = 7560
    End If
    
End Sub

Private Sub panelAcciones_Expand(State As Boolean)
    comprobarAlto
End Sub
Private Sub panelFirmas_Expand(State As Boolean)
    comprobarAlto
End Sub
Private Sub panelModificaciones_Expand(State As Boolean)
    comprobarAlto
End Sub

Private Sub Timer1_Timer()
    cargar_lista_firmas
End Sub

Private Sub TimerAcciones_Timer()
    cargar_lista_acciones
End Sub
