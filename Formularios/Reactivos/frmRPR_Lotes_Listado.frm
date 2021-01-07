VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRPR_Lotes_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Lotes"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13470
   Icon            =   "frmRPR_Lotes_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   13470
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
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
      Height          =   1080
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   13485
      Begin VB.CheckBox chkTodosBotes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10050
         TabIndex        =   8
         Top             =   270
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   825
         Left            =   12105
         Picture         =   "frmRPR_Lotes_Listado.frx":0ABA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1140
      End
      Begin MSDataListLib.DataCombo cmbrpr 
         Height          =   315
         Left            =   1770
         TabIndex        =   9
         Top             =   255
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1770
         TabIndex        =   10
         Top             =   630
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   58195969
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3915
         TabIndex        =   11
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   58195969
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Reactivo"
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
         Left            =   120
         TabIndex        =   14
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricado desde"
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
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   690
         Width           =   1560
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
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
         Index           =   2
         Left            =   3285
         TabIndex        =   12
         Top             =   675
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12285
      Picture         =   "frmRPR_Lotes_Listado.frx":1384
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1185
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3645
      Picture         =   "frmRPR_Lotes_Listado.frx":1C4E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   1185
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2430
      Picture         =   "frmRPR_Lotes_Listado.frx":2518
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1215
      Picture         =   "frmRPR_Lotes_Listado.frx":2DE2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1185
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo lote"
      Height          =   870
      Left            =   0
      Picture         =   "frmRPR_Lotes_Listado.frx":36AC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1185
   End
   Begin VB.CommandButton cmdFicha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficha"
      Height          =   870
      Left            =   4860
      Picture         =   "frmRPR_Lotes_Listado.frx":3F76
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1185
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5055
      Left            =   0
      TabIndex        =   15
      Top             =   1755
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin VB.Label lblmsg 
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
      Left            =   0
      TabIndex        =   17
      Top             =   1440
      Width           =   13485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listado de Lotes"
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
      Height          =   330
      Index           =   4
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   13485
   End
End
Attribute VB_Name = "frmRPR_Lotes_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 50
    Me.Top = 50
    Call cabecera
    'Cargar_Combo cmbrpr, New clsRPR_Tipos
    fdesde = Date - 30
    fhasta = Date
End Sub

Private Sub cmdAnadir_Click()
    frmRPR_Lotes_Detalle.PK = 0
    frmRPR_Lotes_Detalle.Show 1
    cmdBuscar_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    Call buscar
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

Private Sub txtcodigo_GotFocus()
    txtcodigo.BackColor = &H80C0FF
    txtcodigo.SelStart = 0
    txtcodigo.SelLength = Len(txtcodigo)
End Sub
Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtcodigo <> "" Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtcodigo_LostFocus()
    txtcodigo.BackColor = &HFFFFFF
    cargar_codigo (txtcodigo)
    txtcodigo = ""
End Sub

' Funciones auxiliares del formulario
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Número", 1000, lvwColumnCenter
        .Add , , "Fecha creación", 1450, lvwColumnCenter
        .Add , , "Tipo reactivo", 7000, lvwColumnCenter
    End With
End Sub

Public Sub cargar_codigo(CODIGO As String)
    If Trim(txtcodigo) = "" Then
        Exit Sub
    End If
    Dim consulta As String
    Dim rs As ADODB.RecordSet
   On Error GoTo cargar_codigo_Error
    lista.ListItems.Clear
    consulta = "SELECT A.ID_BOTE_PR, " & _
               "       A.NUMERO, " & _
               "       B.CODIGO, " & _
               "       B.NOMBRE, " & _
               "       A.FECHA_FABRICACION, " & _
               "       A.FECHA_CADUCIDAD, " & _
               "       A.VOLUMEN, " & _
               "       C.USUARIO " & _
               " FROM RPR_BOTES A, " & _
               "      RPR_TIPOS B, " & _
               "      USUARIOS C " & _
               " WHERE A.TIPO_REACTIVO_PR_ID = B.ID_TIPO_REACTIVO_PR " & _
               "   AND A.EMPLEADO_ID = C.ID_EMPLEADO " & _
               "   AND A.ID_BOTE_PR = " & CLng(CODIGO)
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        With lista.ListItems.Add(, , Format(rs(0), "00000"))
         .SubItems(1) = Format(rs(1), "00000")
         .SubItems(2) = rs(2)
         .SubItems(3) = rs(3)
         .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
         .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
         .SubItems(6) = rs(6)
         .SubItems(7) = rs(7)
        End With
    End If

   On Error GoTo 0
   Exit Sub

cargar_codigo_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_codigo of Formulario frmRPR_Gestion"
End Sub

Private Sub buscar()
    Dim consulta As String
    Dim strBote As String
    Dim strCaducado As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As ADODB.RecordSet
    Dim f_desde As String
    Dim f_hasta As String
    f_desde = Format(fdesde, "yyyy-mm-dd")
    f_hasta = Format(fhasta, "yyyy-mm-dd")
    Dim IMPORTE As Currency
    IMPORTE = 0
    ' Tipo de Bote
    strBote = ""
    If chkTodosBotes.value = Unchecked Then
        If cmbrpr.Text = "" Then
'            MsgBox "Debe seleccionar un Tipo de Reactivo.", vbExclamation, App.Title
            Exit Sub
        End If
        strBote = " AND A.TIPO_REACTIVO_PR_ID=" & cmbrpr.BoundText
    End If
    ' Fechas
    Dim fecha_desde As String
    fecha_desde = " AND A.FECHA_FABRICACION>='" & f_desde & "'"
    Dim fecha_hasta As String
    fecha_hasta = " AND A.FECHA_FABRICACION<='" & f_hasta & "'"
    ' Query
    consulta = "SELECT A.ID_BOTE_PR, " & _
               "       A.NUMERO, " & _
               "       B.CODIGO, " & _
               "       B.NOMBRE, " & _
               "       A.FECHA_FABRICACION, " & _
               "       A.FECHA_CADUCIDAD, " & _
               "       A.VOLUMEN, " & _
               "       C.USUARIO " & _
               " FROM RPR_BOTES A, " & _
               "      RPR_TIPOS B, " & _
               "      USUARIOS C " & _
               " WHERE A.TIPO_REACTIVO_PR_ID = B.ID_TIPO_REACTIVO_PR " & _
               "   AND A.EMPLEADO_ID = C.ID_EMPLEADO " & _
                          strBote & _
                          fecha_desde & _
                          fecha_hasta & _
                          strCaducado & _
               " ORDER BY A.NUMERO DESC"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , Format(rs(0), "00000"))
                .SubItems(1) = Format(rs(1), "00000")
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
                .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
                .SubItems(6) = rs(6)
                .SubItems(7) = rs(7)
            End With
            rs.MoveNext
        Wend
        lblmsg.Caption = "Botes entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & " (Encontrados : " & rs.RecordCount & ")"
    Else
        lblmsg.Caption = "No existe ningun bote con esos criterios."
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Botes.", vbCritical, Err.Description
End Sub

Public Sub actualizar_lista()
    Dim consulta As String
    Dim rs As ADODB.RecordSet
    consulta = "SELECT A.ID_BOTE_PR, " & _
               "       A.NUMERO, " & _
               "       B.CODIGO, " & _
               "       B.NOMBRE, " & _
               "       A.FECHA_FABRICACION, " & _
               "       A.FECHA_CADUCIDAD, " & _
               "       A.VOLUMEN, " & _
               "       C.USUARIO " & _
               " FROM RPR_BOTES A, " & _
               "      RPR_TIPOS B, " & _
               "      USUARIOS C " & _
               " WHERE A.TIPO_REACTIVO_PR_ID = B.ID_TIPO_REACTIVO_PR " & _
               "   AND A.EMPLEADO_ID = C.ID_EMPLEADO " & _
               "   AND A.ID_BOTE_PR = " & lista.ListItems(lista.SelectedItem.Index)
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = Format(rs(1), "00000")
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = rs(3)
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = Format(rs(4), "dd-mm-yyyy")
        lista.ListItems(lista.SelectedItem.Index).SubItems(5) = Format(rs(5), "dd-mm-yyyy")
        lista.ListItems(lista.SelectedItem.Index).SubItems(6) = rs(6)
        lista.ListItems(lista.SelectedItem.Index).SubItems(7) = rs(7)
    End If
End Sub
' -----------------------------------------
