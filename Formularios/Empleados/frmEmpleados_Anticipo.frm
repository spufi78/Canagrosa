VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpleados_Anticipo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anticipo"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   Icon            =   "frmEmpleados_Anticipo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   885
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5370
      Width           =   1155
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5370
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5370
      Width           =   1155
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5370
      Width           =   1155
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4470
      MaxLength       =   10
      TabIndex        =   0
      Top             =   495
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comentarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   13
      Left            =   45
      TabIndex        =   4
      Top             =   900
      Width           =   5730
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Index           =   1
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   5475
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2820
      Left            =   45
      TabIndex        =   2
      Top             =   2160
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4974
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
   Begin MSComCtl2.DTPicker Fecha 
      Height          =   360
      Left            =   855
      TabIndex        =   3
      Top             =   450
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   14737632
      Format          =   52756481
      CurrentDate     =   38000
      MinDate         =   2
   End
   Begin VB.Label lbltotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1350
      TabIndex        =   8
      Top             =   4995
      Width           =   3210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   3465
      TabIndex        =   7
      Top             =   540
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   45
      TabIndex        =   6
      Top             =   495
      Width           =   660
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Anticipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   15
      TabIndex        =   5
      Top             =   75
      Width           =   5715
   End
End
Attribute VB_Name = "frmEmpleados_Anticipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdAnadir_Click()
    If PK > 0 Then
        insertar_Anticipo
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        PREGUNTA = "Va a ELIMINAR un Anticipo. ¿Esta seguro?"
        If MsgBox(PREGUNTA, vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oc As New clsEmpleados_Anticipo
            If oc.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
                MsgBox "El Anticipo se ha eliminado correctamente.", vbInformation, App.Title
                listado_Anticipos
            End If
        End If
    End If

End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim appword As Word.Application
    Dim docword As Word.Document
    ' Crear copia para su uso
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(App.Path & "\informes\Anticipo.doc")
    ' Cabecera
    Dim tNum2Text As New cNum2Text
    fecha = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
    docword.Tables(1).Rows.Last.Cells(2).Range.Text = fecha_larga(fecha)
    docword.Tables(2).Rows.Last.Cells(1).Range.Text = "CON FECHA DE HOY HE RECIBIDO " & _
            "LA CANTIDAD DE " & UCase(tNum2Text.Numero2Letra(lista.ListItems(lista.SelectedItem.Index).SubItems(2), , 2, "euro", "céntimo", Masculino, Masculino)) & _
            " (" & Format(lista.ListItems(lista.SelectedItem.Index).SubItems(2), "currency") & _
            ") EN CONCEPTO DE ANTICIPO DE MI NOMINA DEL MES DE " & UCase(mes) & "."
    If PK > 0 Then
        Dim operario As New clsEmpleados
        operario.CARGAR (PK)
        docword.Tables(3).Rows.Last.Cells(1).Range.Text = operario.getNOMBRE
    End If
    appword.Visible = True
    Set docword = Nothing
    Set appword = Nothing
    Exit Sub
fallo:
    appword.Quit
    Set docword = Nothing
    Set appword = Nothing
    MsgBox "Se ha producido un error al generar el documento. " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    titulo_ventana
    cabecera_lista
    If PK > 0 Then
        listado_Anticipos
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmEmpleados_Anticipo = Nothing
End Sub
Private Sub lista_Click()
    consulta_anticipo
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
End Sub
Public Sub borrar_campos()
    Dim i As Integer
    For i = 0 To 1
        txtdatos(i) = ""
    Next
End Sub
Public Sub insertar_Anticipo()
    If valida_datos = False Then
        Exit Sub
    End If
    PREGUNTA = "Va a dar de alta un nuevo Anticipo. ¿Esta seguro?"
    If MsgBox(PREGUNTA, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Set oAnticipo = mover_datos
        If oAnticipo.Insertar > 0 Then
            MsgBox "El Anticipo se ha insertado correctamente.", vbInformation, App.Title
            listado_Anticipos
            cmdImprimir_Click
        End If
    End If
End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If txtdatos(0) = "" Then
        MsgBox "La cantidad no puede estar en blanco.", vbCritical, "Error"
        txtdatos(0).SetFocus
        valida_datos = False
        Exit Function
    End If
    If IsNumeric(txtdatos(0)) = False Then
        MsgBox "La cantidad debe ser numérica.", vbCritical, "Error"
        txtdatos(0).SetFocus
        valida_datos = False
        Exit Function
    End If
End Function
Public Function mover_datos() As clsEmpleados_Anticipo
    On Error GoTo fallo
    Dim oexp As New clsEmpleados_Anticipo
    With oexp
        .setEMPLEADO_ID = PK
        .setCANTIDAD = CSng(Replace(txtdatos(0), ".", ","))
        .setCOMENTARIO = txtdatos(1)
        .setFECHA = Format(fecha.value, "yyyy-mm-dd")
    End With
    Set mover_datos = oexp
    Set oexp = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del Anticipo.", vbCritical, Err.Description
End Function
Public Sub listado_Anticipos()
    Dim oexp As New clsEmpleados_Anticipo
    Dim rs As ADODB.RecordSet
    Set rs = oexp.Listado(PK)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("id"))
            .SubItems(1) = Format(rs("fecha"), "dd/mm/yyyy")
            .SubItems(2) = Format(rs("cantidad"), "currency")
            .SubItems(3) = rs("comentario")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    consulta_total
End Sub
Public Sub cabecera_lista()
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Cantidad", 1200, lvwColumnRight)
        .Tag = "Cantidad"
    End With
    With lista.ColumnHeaders.Add(, , "Comentario", 3300, lvwColumnLeft)
        .Tag = "Comentario"
    End With
End Sub
Public Sub titulo_ventana()
    fecha = Date
    If PK > 0 Then
        Dim operario As New clsEmpleados
        operario.CARGAR (PK)
        lbltitulo.Caption = "Anticipos de : " & operario.getNOMBRE
        Me.Caption = lbltitulo.Caption
    End If
End Sub
Public Sub consulta_anticipo()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oexp As New clsEmpleados_Anticipo
    oexp.CARGAR (lista.ListItems(lista.SelectedItem.Index))
    With oexp
        txtdatos(0) = .getCANTIDAD
        txtdatos(1) = .getCOMENTARIO
        fecha = .getFECHA
    End With
    Set oexp = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar el anticipo.", vbCritical, Err.Description
End Sub
Public Sub consulta_total()
    Dim ooa As New clsEmpleados_Anticipo
    lbltotal = "Total anticipos : " & Format(ooa.total(PK), "currency")
End Sub
