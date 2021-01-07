VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmTarifas_Codigos 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Tarifas"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTarifas_Codigos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   11370
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignación de Códigos"
      Height          =   870
      Left            =   8190
      Picture         =   "frmTarifas_Codigos.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7785
      Width           =   1995
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7785
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7785
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7785
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3645
      TabIndex        =   15
      Top             =   8055
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   45
      TabIndex        =   13
      Top             =   720
      Width           =   11265
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   6390
         TabIndex        =   1
         Top             =   210
         Width           =   3675
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1110
         TabIndex        =   0
         Top             =   210
         Width           =   3675
      End
      Begin pryCombo.miCombo cmbfamilia 
         Height          =   330
         Left            =   1110
         TabIndex        =   2
         Top             =   570
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   582
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   285
         Left            =   5640
         TabIndex        =   19
         Top             =   255
         Width           =   630
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   285
         Left            =   90
         TabIndex        =   14
         Top             =   255
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7785
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5325
      Left            =   45
      TabIndex        =   10
      Top             =   1755
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   9393
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   45
      TabIndex        =   16
      Top             =   7110
      Width           =   11265
      Begin VB.TextBox txtdato 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   180
         Width           =   1400
      End
      Begin VB.TextBox txtdato 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1485
         TabIndex        =   4
         Top             =   180
         Width           =   6930
      End
      Begin MSDataListLib.DataCombo cmbfam 
         Height          =   315
         Left            =   8460
         TabIndex        =   5
         Top             =   180
         Width           =   2730
         _ExtentX        =   4815
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
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Asignación de precios a códigos tarifarios"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   390
      Width           =   2925
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10710
      Picture         =   "frmTarifas_Codigos.frx":08D6
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Códigos Tarifarios"
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
      TabIndex        =   11
      Top             =   60
      Width           =   3135
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmTarifas_Codigos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbfamilia_Change()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    If validar Then
        Dim oTC As New clsTarifas_codigos
        With oTC
            .setCODIGO = txtdato(0)
            .setDESCRIPCION = txtdato(1)
            .setFAMILIA_CODIGO_ID = cmbfam.BoundText
            .setPRECIO = moneda_bd("0")
            .Insertar
            cargar_lista
        End With
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar la tarifa. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oTC As New clsTarifas_codigos
            oTC.Eliminar lista.ListItems(lista.SelectedItem.Index)
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If validar Then
        Dim oTC As New clsTarifas_codigos
        With oTC
            .setCODIGO = txtdato(0)
            .setDESCRIPCION = txtdato(1)
            .setFAMILIA_CODIGO_ID = cmbfam.BoundText
            .setPRECIO = moneda_bd("0")
            .Modificar lista.ListItems(lista.SelectedItem.Index)
            cargar_lista
        End With
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    execute_bd "DELETE FROM TARIFAS_CODIGOS"
    execute_bd "DELETE FROM TARIFAS_CODIGOS_FAMILIAS"
    Dim oF As New clsTarifas_codigos_familias
    Dim oc As New clsTarifas_codigos
    
    Dim rs As ADODB.RecordSet
    Dim XLA As Excel.Application
    Dim XLW As Excel.Workbook
    Dim XLS As Excel.Worksheet
    Set XLA = New Excel.Application
    Set XLW = XLA.Workbooks.Open(App.Path & "\TARIFA.xls")
    Set XLS = XLW.Worksheets(1)
    Dim i As Integer
    Dim familia As Integer
    i = 1 ' Primera fila con datos
    While XLS.Cells(i, 2) <> ""
            If Trim(XLS.Cells(i, 1)) = "" Then
                oF.setDESCRIPCION = Trim(XLS.Cells(i, 2))
                familia = oF.Insertar
            Else
                With oc
                    .setCODIGO = XLS.Cells(i, 1)
                    If Trim(XLS.Cells(i, 3)) = "" Or Trim(XLS.Cells(i, 3)) = "N/A" Then
                        .setDESCRIPCION = XLS.Cells(i, 2)
                    Else
                        .setDESCRIPCION = XLS.Cells(i, 2) & " (" & XLS.Cells(i, 3) & ")"
                    End If
                    .setFAMILIA_CODIGO_ID = familia
                    
                    If Trim(XLS.Cells(i, 6)) = "" Then
                        .setPRECIO = moneda_bd("0")
                    Else
                        .setPRECIO = Replace(XLS.Cells(i, 6), ",", ".")
                    End If
                    .Insertar
                End With
            End If
        i = i + 1
    Wend
    XLW.Close
    XLA.Quit
    MsgBox "OK  "
    cargar_lista
End Sub

Private Sub Command2_Click()
    frmTarifas_Codigos_Asignacion.Show 1
End Sub

Private Sub Form_Activate()
    Me.SetFocus
    cargar_lista
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    cargar_botones Me
    cargar_combos
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Código", 1400, lvwColumnCenter
        .Add , , "Descripción", 7000, lvwColumnLeft
        .Add , , "Familia", 2500, lvwColumnCenter
        .Add , , "id_familia", 1, lvwColumnRight
    End With
'    If UCase(usuario.getUSUARIO) = "JULIO" Then
'        Command1.Visible = True
'    End If
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.RecordSet
    Dim oTC As New clsTarifas_codigos
    If cmbfamilia.getTEXTO = "" Then
        Set rs = oTC.Listado(txtfiltro(0), txtfiltro(1), 0)
    Else
        Set rs = oTC.Listado(txtfiltro(0), txtfiltro(1), cmbfamilia.getPK_SALIDA)
    End If
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
'             .SubItems(1) = Format(rs(1), "0000.00")
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
'        lista_Click
    End If
    Set oTC = Nothing
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count <> 0 Then
        txtdato(0) = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        txtdato(1) = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        cmbfam.BoundText = lista.ListItems(lista.SelectedItem.Index).SubItems(4)
'        txtdato(0).SetFocus
    End If
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
        frmTarifas_Codigos_Precios.pk = lista.ListItems(lista.SelectedItem.Index)
        frmTarifas_Codigos_Precios.Show 1
    End If
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub
Private Sub txtdato_GotFocus(Index As Integer)
    txtdato(Index).SelStart = 0
    txtdato(Index).SelLength = Len(txtdato(Index))
    txtdato(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdato_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 46 Then
       KeyAscii = 44
    End If
'    If KeyAscii = 13 Then
'        anadir_precio
'    End If
End Sub
Private Sub txtdato_LostFocus(Index As Integer)
    txtdato(Index).BackColor = vbWhite
End Sub
'Private Sub anadir_precio()
'    If lista.ListItems.Count > 0 Then
'        If txtdato(2) = "" Then
'            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
'            txtdato(2).SetFocus
'            Exit Sub
'        End If
'        If Not IsNumeric(2) Then
'            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
'            txtdato(2).SetFocus
'            Exit Sub
'        End If
'        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = Format(txtdato(2), "currency")
'        Dim oTc As New clsTarifas_codigos
'        oTc.setPRECIO = Replace(txtdato(2), ",", ".")
'        oTc.Modificar_Precio lista.ListItems(lista.SelectedItem.Index)
'        Set oTc = Nothing
'        If lista.ListItems.Count > lista.SelectedItem.Index Then
'            Set lista.SelectedItem = lista.ListItems(lista.SelectedItem.Index + 1)
'            lista.SetFocus
'            lista_Click
'        End If
'    End If
'End Sub
Public Sub cargar_combos()
    llenar_combo cmbfamilia, New clsTarifas_codigos_familias, 0, Me, ""
    cargar_combo cmbfam, New clsTarifas_codigos_familias
End Sub

Public Function validar() As Boolean
    validar = True
    If txtdato(0) = "" Then
        MsgBox "Introduzca un código para la tarifa.", vbInformation, App.Title
        txtdato(0).SetFocus
        validar = False
        Exit Function
    End If
    If txtdato(1) = "" Then
        MsgBox "Introduzca una descripción para la tarifa.", vbInformation, App.Title
        txtdato(1).SetFocus
        validar = False
        Exit Function
    End If
    If cmbfam.Text = "" Then
        MsgBox "Introduzca la familia para la tarifa.", vbInformation, App.Title
        cmbfam.SetFocus
        validar = False
        Exit Function
    
    End If
'        If txtdato(2) = "" Then
'            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
'            txtdato(2).SetFocus
'            validar = False
'            Exit Function
'        End If
'        If Not IsNumeric(2) Then
'            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
'            txtdato(2).SetFocus
'            validar = False
'            Exit Function
'        End If
End Function

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
