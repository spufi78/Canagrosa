VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDescuentos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tipos de Descuento"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   Icon            =   "frmDescuentos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9720
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   30
      TabIndex        =   10
      Top             =   5820
      Width           =   8475
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   885
         Left            =   7230
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdmodificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   885
         Left            =   6030
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdanadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   885
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   3
         Left            =   3990
         TabIndex        =   3
         Top             =   720
         Width           =   765
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   2
         Left            =   2250
         TabIndex        =   2
         Top             =   720
         Width           =   765
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   810
         TabIndex        =   1
         Top             =   720
         Width           =   765
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   0
         Left            =   810
         TabIndex        =   0
         Top             =   300
         Width           =   3945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comercial"
         Height          =   195
         Index           =   3
         Left            =   3180
         TabIndex        =   14
         Top             =   810
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dto. 2"
         Height          =   195
         Index           =   2
         Left            =   1710
         TabIndex        =   13
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dto. 1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6060
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5355
      Left            =   30
      TabIndex        =   8
      Top             =   360
      Width           =   9675
      _ExtentX        =   17066
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
      Caption         =   "Mantenimiento de Tipos de Descuentos"
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
      TabIndex        =   9
      Top             =   30
      Width           =   9660
   End
End
Attribute VB_Name = "frmDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If validar = True Then
        If MsgBox("Va a insertar el tipo de descuento. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDescuento As New clsDescuentos
            With oDescuento
                .setNOMBRE = txtdatos(0)
                .setDTO1 = CSng(txtdatos(1))
                If Trim(txtdatos(2)) = "" Then
                    .setDTO2 = 0
                Else
                    .setDTO2 = CSng(txtdatos(2))
                End If
                .setCOMISION = CSng(txtdatos(3))
                .Insertar
            End With
            cargar_lista
        End If
    End If
    txtdatos(0).SetFocus

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
            Dim oDescuento As New clsDescuentos
            oDescuento.Eliminar (lista.ListItems(lista.SelectedItem.Index).Text)
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
            Dim oDescuento As New clsDescuentos
            With oDescuento
                .setNOMBRE = txtdatos(0)
                .setDTO1 = CSng(txtdatos(1))
                If txtdatos(2) = "" Then
                    .setDTO2 = 0
                Else
                    .setDTO2 = CSng(txtdatos(2))
                End If
                .setCOMISION = CSng(txtdatos(3))
                .Modificar (CInt(lista.ListItems(lista.SelectedItem.Index).Text))
            End With
            cargar_lista
        End If
    End If
    txtdatos(0).SetFocus

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
    With lista.ColumnHeaders.Add(, , "Dto.1", 1400, lvwColumnCenter)
        .Tag = "Dto.1"
    End With
    With lista.ColumnHeaders.Add(, , "Dto.2", 1400, lvwColumnCenter)
        .Tag = "Dto.2"
    End With
    With lista.ColumnHeaders.Add(, , "Comision", 1400, lvwColumnCenter)
        .Tag = "Comision"
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oDescuentos As New clsDescuentos
    Set rs = oDescuentos.Listado
    txtdatos(0) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_descuento"), "00"))
            .SubItems(1) = rs("nombre")
            .SubItems(2) = rs("dto1")
            .SubItems(3) = rs("dto2")
            .SubItems(4) = rs("comision")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    Set oDescuentos = Nothing
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtdatos(0).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        txtdatos(1).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        txtdatos(2).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
        txtdatos(3).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(4)
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
    validar = True
    If txtdatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If txtdatos(1).Text = "" Then
        MsgBox "El dto. 1 no puede estar en blanco.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If txtdatos(3).Text = "" Then
        MsgBox "La comisión no puede estar en blanco.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
End Function

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
    txtdatos(Index).BackColor = &HC0E0FF
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
End Sub
