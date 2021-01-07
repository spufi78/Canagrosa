VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTarifas_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tarifas"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9480
   ClipControls    =   0   'False
   Icon            =   "frmTarifas_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   9480
   Begin VB.CheckBox chkBaja 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar las tarifas dadas de Baja"
      Height          =   285
      Left            =   5175
      TabIndex        =   15
      Top             =   7965
      Width           =   2670
   End
   Begin VB.CommandButton cmdClientes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clientes"
      Height          =   870
      Left            =   3330
      Picture         =   "frmTarifas_Listado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7725
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la tarifa"
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
      Height          =   1305
      Left            =   30
      TabIndex        =   10
      Top             =   6375
      Width           =   9360
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   2
         Left            =   3600
         TabIndex        =   2
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   1
         Left            =   8055
         TabIndex        =   1
         Top             =   945
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   660
         TabIndex        =   0
         Top             =   270
         Width           =   8580
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje a aplicar sobre la tarifa seleccionada"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   765
         Width           =   3390
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje a aplicar sobre la tarifa general"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6255
         TabIndex        =   12
         Top             =   660
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarifa"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7725
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dar de Baja"
      Height          =   870
      Left            =   2235
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7725
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7725
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8340
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7725
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5535
      Left            =   0
      TabIndex        =   7
      Top             =   780
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   9763
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
      Caption         =   "Mantenimiento de Tarifas"
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
      TabIndex        =   9
      Top             =   120
      Width           =   2640
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8805
      Picture         =   "frmTarifas_Listado.frx":1194
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Permite modificar, eliminar y dar de alta tarifas a partir de un porcentaje de la tarifa general"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   420
      Width           =   6315
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   9465
   End
End
Attribute VB_Name = "frmTarifas_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkBaja_Click()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    If validar = True Then
        If MsgBox("Va a insertar la Tarifa. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim porcentaje As String
'E001-I
            Dim porcentaje1 As String, porcentaje2 As String
'E001-F
            If Trim(txtdatos(1).Text) <> "" Then
                porcentaje = Trim(txtdatos(1))
            End If
            If Trim(txtdatos(2).Text) <> "" And porcentaje = "" Then
                If lista.ListItems.Count > 0 Then
                porcentaje = Trim(txtdatos(2))
'E001-I
'                    porcentaje = CSng(txtDatos(2)) + CSng(Replace(Replace(lista.ListItems(lista.SelectedItem.Index).SubItems(1), "%", ""), ".", ","))
'                    porcentaje1 = CSng(txtdatos(2))
'                    porcentaje2 = CSng(Replace(Replace(lista.ListItems(lista.SelectedItem.Index).SubItems(1), "%", ""), ".", ","))
'                    porcentaje = ((porcentaje1 * porcentaje2) / 100) + porcentaje1 + porcentaje2
'E001-F
                Else
                    MsgBox "No existe tarifa sobre la que aplicar.", vbCritical, App.Title
                    Exit Sub
                End If
            End If
            Dim oTarifa As New clsTarifas
            Dim TARIFA As Long
            With oTarifa
                .setNOMBRE = txtdatos(0)
                .setPORCENTAJE = porcentaje
                .setTARIFA_ORIGEN_ID = lista.ListItems(lista.SelectedItem.Index).Text
                .setEN_VIGOR = 1
                TARIFA = .Insertar
            End With
            Dim oTP As New clsTarifas_precios
            Me.MousePointer = 11
            oTP.Crear_Tarifa_Nueva TARIFA, lista.ListItems(lista.SelectedItem.Index).Text, porcentaje
            MsgBox "La tarifa se ha generado correctamente.", vbInformation + vbOKOnly, App.Title
            Me.MousePointer = 0
            cargar_lista
        End If
    End If
    txtdatos(0).SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClientes_Click()
    If lista.ListItems.Count > 0 Then
        frmTarifas_Clientes.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmTarifas_Clientes.Show 1
'        If MsgBox("¿Refrescar lista?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'            cargar_lista
'        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la tarifa. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oTarifa As New clsTarifas
            oTarifa.Eliminar (lista.ListItems(lista.SelectedItem.Index).Text)
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtdatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("La modificación solo modifica el nombre. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oTarifa As New clsTarifas
            oTarifa.setNOMBRE = txtdatos(0)
            oTarifa.Modificar (lista.ListItems(lista.SelectedItem.Index).Text)
            cargar_lista
            borrar_campos
        End If
    End If
    txtdatos(0).SetFocus
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 200
    Me.Left = 200
    cargar_botones Me
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Tarifa", 3400, lvwColumnLeft
        .Add , , "Tarifa Origen", 3400, lvwColumnLeft
        .Add , , "% s.Origen", 1200, lvwColumnCenter
        .Add , , "Clientes", 1000, lvwColumnCenter
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.RecordSet
    Dim oTarifa As New clsTarifas
    Set rs = oTarifa.Listado
    borrar_campos
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            If chkBaja.value = Checked Or (chkBaja.value = Unchecked And rs(4) = 1) Then
                With lista.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
                 .SubItems(2) = rs(2)
                 .SubItems(3) = rs(3) & " %"
                 .SubItems(4) = rs(5)
                End With
                If rs(4) = 0 Then ' Si esta de baja, ponerla en rojo
                    colorear_linea lista.ListItems.Count
                End If
            End If
            
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSoluciones = Nothing
    Set rs = Nothing
End Sub
Private Sub colorear_linea(fila As Integer)
    Dim i As Integer
    Dim color As Long
    color = vbRed
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
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
        txtdatos(0).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
    End If
End Sub
Public Sub borrar_campos()
    txtdatos(0) = ""
    txtdatos(1) = ""
    txtdatos(2) = ""
End Sub

Private Function validar() As Boolean
    validar = True
    If txtdatos(0).Text = "" Then
        MsgBox "La descripción de la tarifa no puede estar en blanco.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtdatos(1).Text) = "" And Trim(txtdatos(2).Text) = "" Then
        MsgBox "Inserte algún porcentaje para crear la tarifa (0 para igualarla)", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtdatos(1).Text) <> "" And Trim(txtdatos(2).Text) <> "" Then
        MsgBox "No puede seleccionar los dos porcentajes al mismo tiempo.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtdatos(1).Text) <> "" And Not IsNumeric(txtdatos(1).Text) Then
        MsgBox "El porcentaje debe ser numérico.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtdatos(2).Text) <> "" And Not IsNumeric(txtdatos(2).Text) Then
        MsgBox "El porcentaje debe ser numérico.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
End Function

Private Sub lista_DblClick()
    cmdClientes_Click
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 46 And Index > 0 Then
       KeyAscii = 44
    End If
End Sub
