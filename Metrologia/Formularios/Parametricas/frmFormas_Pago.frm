VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormas_Pago 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Formas de Pago"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   Icon            =   "frmFormas_Pago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   60
      TabIndex        =   13
      Top             =   5760
      Width           =   8895
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         Left            =   3480
         TabIndex        =   6
         Top             =   1170
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CheckBox chkRecibos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Genera Recibos (Efectos)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   1215
         Width           =   2205
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   3990
         TabIndex        =   4
         Top             =   885
         Width           =   765
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   1140
         TabIndex        =   3
         Top             =   855
         Width           =   765
      End
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   885
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   435
         Width           =   1155
      End
      Begin VB.CommandButton cmdmodificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   885
         Left            =   6450
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   435
         Width           =   1155
      End
      Begin VB.CommandButton cmdanadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   885
         Left            =   5250
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   435
         Width           =   1155
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   3990
         TabIndex        =   2
         Top             =   570
         Width           =   765
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   1
         Top             =   540
         Width           =   765
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   180
         Width           =   4200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Forma"
         Height          =   195
         Index           =   5
         Left            =   2445
         TabIndex        =   19
         Top             =   1215
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aplazamiento 1º Recibo"
         Height          =   195
         Index           =   4
         Left            =   2220
         TabIndex        =   18
         Top             =   915
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Día Cobro"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   915
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dias entre Recibos"
         Height          =   195
         Index           =   2
         Left            =   2550
         TabIndex        =   16
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vencimientos"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6195
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5355
      Left            =   30
      TabIndex        =   11
      Top             =   360
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9446
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14609914
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
      Caption         =   "Mantenimiento de Formas de Pago"
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
      Height          =   315
      Index           =   3
      Left            =   30
      TabIndex        =   12
      Top             =   30
      Width           =   10215
   End
End
Attribute VB_Name = "frmFormas_Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If validar = True Then
        If MsgBox("Va a insertar la forma de pago. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oFp As New clsForma_pago
            With oFp
                .setNOMBRE = txtDatos(0)
                .setVENCIMIENTOS = CInt(txtDatos(1))
                .setDIAS = CInt(txtDatos(2))
                .setDIA_COBRO = CInt(txtDatos(3))
                .setAPLAZAMIENTO = CInt(txtDatos(4))
                .setRECIBOS = chkRecibos.Value
                .setTIPO_FORMA = txtDatos(5)
                If .Insertar > 0 Then
                    cargar_lista
                End If
            End With
        End If
    End If
    txtDatos(0).SetFocus

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmFormas_Pago"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error
   If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR el tipo de descuento. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oFp As New clsForma_pago
            oFp.Eliminar (lista.ListItems(lista.SelectedItem.Index).Text)
            cargar_lista
        End If
   End If
   On Error GoTo 0
   Exit Sub
cmdEliminar_Click_Error:
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminar_Click of Formulario frmFormas_Pago"
End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error

    If validar = True Then
        If MsgBox("Va a modificar el tipo de descuento. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oFp As New clsForma_pago
            With oFp
                .setNOMBRE = txtDatos(0)
                .setVENCIMIENTOS = CInt(txtDatos(1))
                .setDIAS = CInt(txtDatos(2))
                .setDIA_COBRO = CInt(txtDatos(3))
                .setAPLAZAMIENTO = CInt(txtDatos(4))
                .setRECIBOS = chkRecibos.Value
                .setTIPO_FORMA = txtDatos(5)
                .Modificar (CInt(lista.ListItems(lista.SelectedItem.Index).Text))
            End With
            actualizar_lista
'            cargar_lista
        End If
    End If
    txtDatos(0).SetFocus

   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmFormas_Pago"
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "ID", 700, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 4500, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Vencimientos", 900, lvwColumnCenter)
        .Tag = "Vencimientos"
    End With
    With lista.ColumnHeaders.Add(, , "Dias", 900, lvwColumnCenter)
        .Tag = "Dias"
    End With
    With lista.ColumnHeaders.Add(, , "Dia Cobro", 900, lvwColumnCenter)
        .Tag = "Dia Cobro"
    End With
    With lista.ColumnHeaders.Add(, , "Aplazamiento", 900, lvwColumnCenter)
        .Tag = "Aplazamiento"
    End With
    With lista.ColumnHeaders.Add(, , "Efectos", 900, lvwColumnCenter)
        .Tag = "Recibos"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo", 0, lvwColumnCenter)
        .Tag = "Tipo"
    End With
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oFp As New clsForma_pago
    Set rs = oFp.Listado
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_forma_pago"), "000"))
            .SubItems(1) = rs("nombre")
            If IsNull(rs("vencimientos")) Then
                .SubItems(2) = "0"
            Else
                .SubItems(2) = rs("vencimientos")
            End If
            If IsNull(rs("dias")) Then
                .SubItems(3) = "0"
            Else
                .SubItems(3) = rs("dias")
            End If
            .SubItems(4) = rs("dia_cobro")
            .SubItems(5) = rs("aplazamiento")
            If rs("recibos") = 1 Then
                .SubItems(6) = "Si"
            Else
                .SubItems(6) = "No"
            End If
            .SubItems(7) = rs("TIPO_FORMA")
           End With
           rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    Set oFp = Nothing
    Set rs = Nothing
End Sub
Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim oFp As New clsForma_pago
    Set rs = oFp.Listado_ID(lista.ListItems(lista.SelectedItem.Index).Text)
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems(lista.SelectedItem.Index)
            .SubItems(1) = rs("nombre")
            If IsNull(rs("vencimientos")) Then
                .SubItems(2) = "0"
            Else
                .SubItems(2) = rs("vencimientos")
            End If
            If IsNull(rs("dias")) Then
                .SubItems(3) = "0"
            Else
                .SubItems(3) = rs("dias")
            End If
            .SubItems(4) = rs("dia_cobro")
            .SubItems(5) = rs("aplazamiento")
            If rs("recibos") = 1 Then
                .SubItems(6) = "Si"
            Else
                .SubItems(6) = "No"
            End If
            .SubItems(7) = rs("TIPO_FORMA")
           End With
           rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    Set oFp = Nothing
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        txtDatos(1).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        txtDatos(2).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
        txtDatos(3).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(4)
        txtDatos(4).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(5)
        If lista.ListItems(lista.SelectedItem.Index).SubItems(6) = "Si" Then
            chkRecibos.Value = Checked
        Else
            chkRecibos.Value = Unchecked
        End If
        txtDatos(5).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
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
Public Function validar() As Boolean
    validar = False
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
        Exit Function
    End If
    If txtDatos(1).Text = "" Then
        MsgBox "El vencimiento puede estar en blanco.", vbCritical, App.Title
        Exit Function
    End If
    If IsNumeric(txtDatos(1)) = False Then
        MsgBox "El vencimiento tiene que ser numérico.", vbCritical, App.Title
        Exit Function
    End If
    If txtDatos(2).Text = "" Then
        MsgBox "Los días no pueden estar en blanco.", vbCritical, App.Title
        Exit Function
    End If
    If IsNumeric(txtDatos(2)) = False Then
        MsgBox "El dias tienen que ser numérico.", vbCritical, App.Title
        Exit Function
    End If
    If txtDatos(3).Text = "" Then
        MsgBox "El día de cobro no pueden estar en blanco.", vbCritical, App.Title
        Exit Function
    End If
    If IsNumeric(txtDatos(3)) = False Then
        MsgBox "El dia de cobro tienen que ser numérico.", vbCritical, App.Title
        Exit Function
    End If
    If txtDatos(4).Text = "" Then
        MsgBox "El aplazamiento no pueden estar en blanco.", vbCritical, App.Title
        Exit Function
    End If
    If IsNumeric(txtDatos(4)) = False Then
        MsgBox "El aplazamiento tiene que ser numérico.", vbCritical, App.Title
        Exit Function
    End If
    If chkRecibos.Value = Checked Then
        If txtDatos(1) = 0 Then
            MsgBox "El número de recibos no puede ser cero.", vbCritical, App.Title
            txtDatos(1).SetFocus
            Exit Function
        End If
        If txtDatos(2) = 0 Then
            MsgBox "Los días entre vencimientos no puede ser cero.", vbCritical, App.Title
            txtDatos(2).SetFocus
            Exit Function
        End If
        If txtDatos(4) = 0 Then
            MsgBox "El aplazamiento del 1º recibo no puede ser cero.", vbCritical, App.Title
            txtDatos(4).SetFocus
            Exit Function
        End If
            
    End If
    validar = True
End Function

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
    txtDatos(Index).BackColor = &HC0E0FF
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
