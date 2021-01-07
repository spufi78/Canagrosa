VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmPDF 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cliente de Etiquetado"
   ClientHeight    =   9060
   ClientLeft      =   5460
   ClientTop       =   3900
   ClientWidth     =   8055
   DrawWidth       =   10
   Icon            =   "frmPDF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Eliminar Etiquetas"
      Height          =   645
      Left            =   90
      Picture         =   "frmPDF.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5985
      Width           =   1770
   End
   Begin VB.CommandButton cmdMinimizar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Minimizar"
      Height          =   645
      Left            =   6210
      Picture         =   "frmPDF.frx":6B5C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6030
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2340
      Left            =   90
      TabIndex        =   5
      Top             =   6660
      Visible         =   0   'False
      Width           =   6045
      Begin VB.TextBox txtImpresoras 
         Height          =   1275
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   900
         Width           =   5730
      End
      Begin VB.CommandButton cmdInsertar 
         Caption         =   "Insertar"
         Height          =   285
         Left            =   3060
         TabIndex        =   3
         Top             =   450
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   135
         TabIndex        =   0
         Top             =   450
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2295
         TabIndex        =   2
         Top             =   450
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1215
         TabIndex        =   1
         Top             =   450
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   16
         Left            =   2385
         TabIndex        =   13
         Top             =   270
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   15
         Left            =   1530
         TabIndex        =   12
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   405
         TabIndex        =   11
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "3."
         Height          =   195
         Index           =   4
         Left            =   4995
         TabIndex        =   9
         Top             =   615
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2.Muestras"
         Height          =   195
         Index           =   1
         Left            =   4995
         TabIndex        =   7
         Top             =   405
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1.REX"
         Height          =   195
         Index           =   0
         Left            =   4995
         TabIndex        =   6
         Top             =   180
         Width           =   465
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5325
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9393
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6525
      Top             =   5805
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LABORATORIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   45
      TabIndex        =   14
      Top             =   0
      Width           =   7935
   End
   Begin XtremeSuiteControls.TrayIcon TrayIcon1 
      Left            =   6210
      Top             =   5805
      _Version        =   851970
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   16
      Text            =   "GESLAB : Generador de Informes v2.0"
      Picture         =   "frmPDF.frx":D3AE
   End
   Begin VB.Label tot 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   8
      Top             =   6705
      Width           =   1230
   End
   Begin VB.Menu opMenu 
      Caption         =   "Menu"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu opRestaurar 
         Caption         =   "Restaurar"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Minimized As Boolean
Private Sub cmdEliminar_Click()
    If MsgBox("¿Esta seguro de borrar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oEtiquetas As New clsEtiquetas
        oEtiquetas.EliminarCentro ReadINI(App.Path + "\config.ini", "Otros", "CENTRO_ID")
        Set oEtiquetas = Nothing
        cargar_lista
    End If
End Sub

Private Sub cmdInsertar_Click()
   On Error GoTo cmdInsertar_Click_Error

    On Error Resume Next
    Dim oEtiquetas As New clsEtiquetas
    If Text1 <> "" And Text2 <> "" And Text3 <> "" Then
        Timer1.Enabled = False
        Dim I As Long
        For I = CLng(Text1) To CLng(Text3)
            With oEtiquetas
                .setCENTRO_ID = ReadINI(App.Path + "\config.ini", "Otros", "CENTRO_ID")
                .setTIPO_ID = Text2
                .setID = I
                .setUSUARIO_ID = 18
                .Insertar
            End With
        Next
        Timer1.Enabled = True
    End If

   On Error GoTo 0
   Exit Sub

cmdInsertar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdInsertar_Click of Formulario frmPDF"
End Sub

Private Sub cmdMinimizar_Click()
    If Not Minimized Then
        TrayIcon1.MinimizeToTray Me.Hwnd
        Minimized = True
    Else
        TrayIcon1.MaximizeFromTray Me.Hwnd
        Minimized = False
    End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Then ' Tecla F11
        Frame1.Visible = Not Frame1.Visible
        If Frame1.Visible = True Then
            Me.Height = 9435
            Text1.SetFocus
        Else
            Me.Height = 7365
        End If
    End If
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    Me.Caption = Me.Caption
    Me.Height = 7365
    cabecera
    cargar_impresoras
    cargar_lista

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmPDF"
End Sub
Private Sub cargar_impresoras()
    Dim IMPRESORA As Printer
   On Error GoTo cargar_impresoras_Error
    Dim encontrada As Boolean
    encontrada = False
    For Each prnPrinter In Printers
        If UCase(prnPrinter.DeviceName) = UCase(Trim(ReadINI(App.Path + "\config.ini", "Otros", "IMPRESORA_G"))) Then
            txtImpresoras.Text = txtImpresoras.Text & " ===> " & prnPrinter.DeviceName & " <=== " & vbNewLine
            encontrada = True
        Else
            txtImpresoras.Text = txtImpresoras.Text & prnPrinter.DeviceName & vbNewLine
        End If
    Next
    If Not encontrada Then
        MsgBox "No se ha encontrado instalada la impresora del parámetro de impresora GRANDE.", vbCritical, App.Title
    End If
    
    encontrada = False
    For Each prnPrinter In Printers
        If UCase(prnPrinter.DeviceName) = UCase(Trim(ReadINI(App.Path + "\config.ini", "Otros", "IMPRESORA_P"))) Then
            txtImpresoras.Text = txtImpresoras.Text & " ===> " & prnPrinter.DeviceName & " <=== " & vbNewLine
            encontrada = True
        Else
            txtImpresoras.Text = txtImpresoras.Text & prnPrinter.DeviceName & vbNewLine
        End If
    Next
    If Not encontrada Then
        MsgBox "No se ha encontrado instalada la impresora del parámetro de impresora PEQUEÑA.", vbCritical, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cargar_impresoras_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_impresoras of Formulario frmPDF"

End Sub
Private Sub cabecera()
   On Error GoTo cabecera_Error

    With Lista.ColumnHeaders
        .Add , , "ID", 500, lvwColumnLeft
        .Add , , "Tipo", 1500, lvwColumnCenter
        .Add , , "Identificador", 1200, lvwColumnCenter
        .Add , , "Usuario", 1800, lvwColumnCenter
        .Add , , "Estado", 700, lvwColumnCenter
        .Add , , "Fecha", 1800, lvwColumnCenter
    End With
    Dim oLab As New clsCentros
    oLab.Carga ReadINI(App.Path + "\config.ini", "Otros", "CENTRO_ID")
    lblTitulo = oLab.getNOMBRE
    Set oLab = Nothing

   On Error GoTo 0
   Exit Sub

cabecera_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cabecera of Formulario frmPDF"
End Sub

Public Sub cargar_lista()
    On Error GoTo fallo
    Dim oEtiquetas As New clsEtiquetas
    Dim rs As ADODB.Recordset
    Set rs = oEtiquetas.Listado(ReadINI(App.Path + "\config.ini", "Otros", "CENTRO_ID"))
    Lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With Lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
            End With
            rs.MoveNext
        Loop Until rs.EOF
        DoEvents
    End If
    Exit Sub
fallo:
    MsgBox "Error al recuperar los datos de la lista.", vbCritical, App.Title
End Sub

Private Sub ib_Menu()
    On Error Resume Next
    PopupMenu opMenu(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub opMenu_Click(Index As Integer)
    Me.Visible = True
End Sub
Private Sub Text1_LostFocus()
    Text3 = Text1
End Sub

Private Sub Timer1_Timer()
    DoEvents
    Dim I As Integer
    I = 1
    Dim CARGAR As Boolean
    CARGAR = True
    While I < Lista.ListItems.Count And CARGAR = True
        If Lista.ListItems(I).SubItems(4) = 0 Then
            CARGAR = False
        End If
        I = I + 1
    Wend
    If CARGAR Then
        cargar_lista
    End If
    imprimir
End Sub

Public Sub imprimir()
    On Error Resume Next
    Dim I As Integer
    Dim IMPRESORA As Integer
    tot = "Total : " & Lista.ListItems.Count
    For I = 1 To Lista.ListItems.Count
        If Lista.ListItems(I).SubItems(4) = 0 Then ' ESTADO
          Timer1.Enabled = False
          Dim oEtiqueta As New clsEtiquetas
          Lista.ListItems(I).SubItems(4) = 1 ' ESTADO : IMPRIMIENDO
          If generar_etiqueta(CLng(Lista.ListItems(I).Text)) = True Then
            oEtiqueta.Impreso Lista.ListItems(I).Text
            Lista.ListItems.Remove I
          Else
            oEtiqueta.ERROR Lista.ListItems(I).Text
            Lista.ListItems(I).SubItems(4) = 3 ' ESTADO : ERROR
          End If
          DoEvents
          Timer1.Enabled = True
          Timer1_Timer
          Exit Sub
        End If
    Next
End Sub
Private Sub TrayIcon1_DblClick()
    If (Minimized) Then cmdMinimizar_Click
End Sub

Private Function generar_etiqueta(ID_ETIQUETA As Long) As Boolean
    Dim res As Boolean
    res = False
    Dim oEtiqueta As New clsEtiquetas
    With oEtiqueta
        .CARGAR ID_ETIQUETA
        Dim impresoraAnterior As String
        impresoraAnterior = Impresora_Predeterminada
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("select carpeta,informe,tamano from etiquetas_rpt where id_tipo = " & .getTIPO_ID)
        
        Establecer_Impresora Trim(ReadINI(App.Path + "\config.ini", "Otros", "IMPRESORA_" & rs(2)))
        
        Select Case .getTIPO_ID
            Case ETIQUETAS_TIPOS.ETIQUETAS_TIPOS_REX  ' REX
                etiqueta_REACTIVO CStr(.getID), .getUSUARIO_ID, rs("carpeta"), rs("informe")
                res = True
            Case ETIQUETAS_TIPOS.ETIQUETAS_TIPOS_MUESTRAS ' MUESTRAS
                etiqueta_MUESTRA CStr(.getID), .getUSUARIO_ID, rs("carpeta"), rs("informe")
                res = True
            Case ETIQUETAS_TIPOS.ETIQUETAS_TIPOS_EQUIPOS_CAL    ' CAL
                etiqueta_CALIBRACION CStr(.getID), .getUSUARIO_ID, rs("carpeta"), rs("informe")
                res = True
            Case ETIQUETAS_TIPOS.ETIQUETAS_TIPOS_EQUIPOS_VER    ' VER
                etiqueta_VERIFICACION CStr(.getID), .getUSUARIO_ID, rs("carpeta"), rs("informe")
                res = True
            Case ETIQUETAS_TIPOS.ETIQUETAS_TIPOS_EQUIPOS     ' EQUIPOS
                etiqueta_EQUIPO CStr(.getID), .getUSUARIO_ID, rs("carpeta"), rs("informe")
                res = True
            Case ETIQUETAS_TIPOS.ETIQUETAS_TIPOS_RPR  ' RPR
                etiqueta_RPR CStr(.getID), .getUSUARIO_ID, rs("carpeta"), rs("informe")
                res = True
        End Select
        Establecer_Impresora impresoraAnterior
    End With
    Set oEtiqueta = Nothing
    generar_etiqueta = res
End Function
