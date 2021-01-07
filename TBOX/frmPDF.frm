VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmPDF 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Servidor de Impresión"
   ClientHeight    =   8820
   ClientLeft      =   5475
   ClientTop       =   3915
   ClientWidth     =   11355
   DrawWidth       =   10
   Icon            =   "frmPDF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstActive 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1620
      Left            =   9900
      TabIndex        =   33
      Top             =   6210
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generación"
      Height          =   2880
      Left            =   45
      TabIndex        =   7
      Top             =   5895
      Visible         =   0   'False
      Width           =   9825
      Begin VB.Frame Frame2 
         Caption         =   "PDF"
         Height          =   600
         Left            =   135
         TabIndex        =   26
         Top             =   2115
         Width           =   8520
         Begin VB.CommandButton cmdDesprotegerPDF 
            Caption         =   "DesProteger"
            Height          =   285
            Left            =   7380
            TabIndex        =   30
            Top             =   180
            Width           =   1050
         End
         Begin VB.CommandButton cmdprotegerPDF 
            Caption         =   "Proteger"
            Height          =   285
            Left            =   6435
            TabIndex        =   29
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox txtruta 
            Height          =   285
            Left            =   630
            TabIndex        =   27
            Top             =   180
            Width           =   5730
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ruta"
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   28
            Top             =   225
            Width           =   345
         End
      End
      Begin VB.CommandButton cmdReConsolidar 
         Caption         =   "Re-Consolidar Baños"
         Height          =   480
         Left            =   135
         TabIndex        =   19
         Top             =   1620
         Width           =   1245
      End
      Begin VB.CheckBox chkFecha 
         Caption         =   "Generar con fecha de cierre"
         Height          =   465
         Left            =   7065
         TabIndex        =   17
         Top             =   180
         Width           =   1515
      End
      Begin VB.CommandButton cmdedicion 
         Caption         =   "Poner Ed.0"
         Height          =   240
         Left            =   3870
         TabIndex        =   16
         Top             =   1170
         Width           =   915
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   285
         Left            =   3870
         TabIndex        =   15
         Top             =   855
         Width           =   915
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "Query"
         Height          =   285
         Left            =   3870
         TabIndex        =   9
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox Text4 
         Height          =   1050
         Left            =   135
         TabIndex        =   8
         Text            =   "select * from muestras where"
         Top             =   540
         Width           =   3660
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Insertar"
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   225
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   135
         TabIndex        =   0
         Top             =   240
         Width           =   1050
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2295
         TabIndex        =   2
         Top             =   225
         Width           =   555
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Re"
         Height          =   285
         Left            =   3870
         TabIndex        =   4
         Top             =   225
         Width           =   915
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1215
         TabIndex        =   1
         Top             =   225
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "70. Firma Informe"
         Height          =   195
         Index           =   15
         Left            =   7785
         TabIndex        =   35
         Top             =   1755
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "65. Firma Factura"
         Height          =   195
         Index           =   14
         Left            =   7785
         TabIndex        =   34
         Top             =   1530
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "56. Imprimir Norma"
         Height          =   195
         Index           =   13
         Left            =   7785
         TabIndex        =   32
         Top             =   1305
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "51. Proteger Norma"
         Height          =   195
         Index           =   12
         Left            =   7785
         TabIndex        =   31
         Top             =   855
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "55. Imprimir documento"
         Height          =   195
         Index           =   10
         Left            =   7785
         TabIndex        =   25
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "50. Proteger Calidad"
         Height          =   195
         Index           =   9
         Left            =   7785
         TabIndex        =   24
         Top             =   630
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "60. Consolidación Excel"
         Height          =   195
         Index           =   8
         Left            =   4905
         TabIndex        =   23
         Top             =   1935
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "41. Convertir a PDF (Ruta, Destino)"
         Height          =   195
         Index           =   7
         Left            =   4905
         TabIndex        =   22
         Top             =   1713
         Width           =   2505
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "40. PNT"
         Height          =   195
         Index           =   6
         Left            =   4905
         TabIndex        =   21
         Top             =   1494
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "20. Alodine"
         Height          =   195
         Index           =   5
         Left            =   4905
         TabIndex        =   20
         Top             =   1275
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "3.Previsualización"
         Height          =   195
         Index           =   4
         Left            =   4905
         TabIndex        =   18
         Top             =   618
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "11.Recepcion + Imp."
         Height          =   195
         Index           =   3
         Left            =   4905
         TabIndex        =   13
         Top             =   1056
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "10.Recepcion"
         Height          =   195
         Index           =   2
         Left            =   4905
         TabIndex        =   12
         Top             =   837
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2.Reimpresion"
         Height          =   195
         Index           =   1
         Left            =   4905
         TabIndex        =   11
         Top             =   399
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1.Informe"
         Height          =   195
         Index           =   0
         Left            =   4905
         TabIndex        =   10
         Top             =   180
         Width           =   660
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Minimizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9990
      TabIndex        =   5
      Top             =   8145
      Width           =   1185
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5775
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   10186
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
      Left            =   9855
      Top             =   6390
   End
   Begin XtremeSuiteControls.TrayIcon TrayIcon1 
      Left            =   9900
      Top             =   7920
      _Version        =   851970
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   16
      Text            =   "GESLAB : Generador de Informes v2.0"
      Picture         =   "frmPDF.frx":030A
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
      Left            =   10080
      TabIndex        =   14
      Top             =   5850
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

Private Sub cmdClear_Click()
    If MsgBox("¿Esta seguro de borrar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        execute_bd "delete from impresion"
    End If
End Sub

Private Sub cmdDesprotegerPDF_Click()
    If txtruta <> "" Then
        Timer1.Enabled = False
        Dim oImp As New clsImpresion
        With oImp
            oImp.setMUESTRA_ID = 0
            oImp.setTIPO = 51
            oImp.setEMPLEADO_ID = 18
            oImp.setRUTA_ORIGEN = Replace(txtruta, "\", "/")
            oImp.Insertar
        End With
        Timer1.Enabled = True
    Else
        MsgBox "Inserte la ruta.", vbExclamation, App.Title
    End If
End Sub

Private Sub cmdedicion_Click()
    On Error Resume Next
    Dim oImp As New clsImpresion
    If Text1 <> "" And Text3 <> "" Then
        Dim i As Long
        Dim CONSULTA As String
        For i = CLng(Text1) To CLng(Text3)
            CONSULTA = "update muestras set ult_edicion_imp = 0 where id_muestra = " & i
            execute_bd CONSULTA
        Next
    End If
    MsgBox "Ok."
End Sub

Private Sub cmdprotegerPDF_Click()
    If txtruta <> "" Then
        Timer1.Enabled = False

        Dim oImp As New clsImpresion
        With oImp
            oImp.setMUESTRA_ID = 0
            oImp.setTIPO = 50
            oImp.setEMPLEADO_ID = 18
            oImp.setRUTA_ORIGEN = Replace(txtruta, "\", "/")
            oImp.Insertar
        End With
        Timer1.Enabled = True
    Else
        MsgBox "Inserte la ruta.", vbExclamation, App.Title
    End If
End Sub

Private Sub cmdReConsolidar_Click()
    On Error GoTo error_consolidar
    Dim oImp As New clsImpresion
    If Text1 <> "" Then
        Timer1.Enabled = False
        oImp.setMUESTRA_ID = CLng(Text1.Text)
        oImp.setTIPO = 60
        oImp.setEMPLEADO_ID = 1
        oImp.Insertar
        Timer1.Enabled = True
    End If
    
    Set oImp = Nothing
    
    Exit Sub
error_consolidar:
MsgBox "Debe señalar el ID del Baño en el primer Texto"
Set oImp = Nothing
End Sub

Private Sub cmdQuery_Click()
    If Trim(Text2) = "" Then
        MsgBox "Inserte el tipo", vbInformation, App.Title
        Text2.SetFocus
        Exit Sub
    End If
    Dim oImp As New clsImpresion
    Dim RS As ADODB.Recordset
    Dim c As String
    Timer1.Enabled = False
    c = Text4
    Set RS = datos_bd(c)
    If RS.RecordCount <> 0 Then
        Do
            oImp.setMUESTRA_ID = RS(0)
            oImp.setTIPO = CInt(Text2)
            oImp.setEMPLEADO_ID = 1
            oImp.Insertar
            RS.MoveNext
        Loop Until RS.EOF
    End If
    Timer1.Enabled = True
End Sub
Private Sub Command1_Click()
    If Not Minimized Then
        TrayIcon1.MinimizeToTray Me.hWnd
        Minimized = True
    Else
        TrayIcon1.MaximizeFromTray Me.hWnd
        Minimized = False
    End If
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim oImp As New clsImpresion
    If Text1 <> "" And Text2 <> "" And Text3 <> "" Then
        Timer1.Enabled = False
        Dim i As Long
        For i = CLng(Text1) To CLng(Text3)
            If chkprimera.Value = 1 Then
                execute_bd "update muestras set ult_edicion_impresa = 0 where id_muestra = " & i
            End If
            oImp.setMUESTRA_ID = i
            oImp.setTIPO = Text2
            oImp.setEMPLEADO_ID = 1
            oImp.Insertar
        Next
        Timer1.Enabled = True
    End If
End Sub

Private Sub Command3_Click()
    If Lista.ListItems.Count > 0 Then
        Dim oImp As New clsImpresion
        oImp.setMUESTRA_ID = Lista.ListItems(Lista.SelectedItem.Index).Text
        oImp.setTIPO = Lista.ListItems(Lista.SelectedItem.Index).SubItems(1)
        oImp.setEMPLEADO_ID = 18
        oImp.Insertar
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Then ' Tecla F11
        Frame1.Visible = Not Frame1.Visible
        If Frame1.Visible = True Then
            Text1.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    If CrearConexionGlobal = False Then
        MsgBox "Error al crear la conexión global. Contacte con mantenimiento.", vbCritical, App.Title
        End
    End If
    Me.Caption = Me.Caption & " (Host: " & ReadINI(App.Path + "\config.ini", "server", "ip") & " -> BD: " & database & ")"
'    ib.mensaje = Me.Caption
    cabecera
    cargar_lista
End Sub

Public Sub cabecera()
    With Lista.ColumnHeaders
        .Add , , "Identif.", 800, lvwColumnLeft
        .Add , , "Tipo", 600, lvwColumnCenter
        .Add , , "Usuario", 1000, lvwColumnCenter
        .Add , , "Puesto", 1700, lvwColumnCenter
        .Add , , "Estado", 700, lvwColumnCenter
        .Add , , "Fecha", 1000, lvwColumnCenter
        .Add , , "Hora", 1000, lvwColumnCenter
        .Add , , "ID_EMPLEADO", 1, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "RUTA_ORIGEN", 1500, lvwColumnLeft
        .Add , , "RUTA_DESTINO", 1500, lvwColumnLeft
    End With
End Sub

Public Sub cargar_lista()
    On Error GoTo fallo
    Dim oImpresion As New clsImpresion
    Dim RS As ADODB.Recordset
    Set RS = oImpresion.Listado
    Lista.ListItems.Clear
    If RS.RecordCount <> 0 Then
        Do
            With Lista.ListItems.Add(, , RS(0))
             .SubItems(1) = RS(1)
             .SubItems(2) = UCase(RS(2))
             .SubItems(3) = RS(3)
             .SubItems(4) = RS(4)
             .SubItems(5) = RS(5)
             .SubItems(6) = RS(6)
             .SubItems(7) = RS(7)
             .SubItems(8) = RS(8)
             .SubItems(9) = RS("RUTA_ORIGEN")
             .SubItems(10) = RS("RUTA_DESTINO")
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
    Exit Sub
fallo:
    enviar_correo "julio.gonzalez@ixitec.net", "", "", False, "", "ERROR SERVIDOR IMPRESION : " & Err.Description
    MsgBox "Error al recuperar los datos de la lista.", vbCritical, App.Title
End Sub

Private Sub ib_Menu()
    On Error Resume Next
    PopupMenu opMenu(0)
End Sub
Private Sub opMenu_Click(Index As Integer)
    Me.Visible = True
End Sub
Private Sub Text1_LostFocus()
    Text3 = Text1
End Sub

Private Sub Timer1_Timer()
    DoEvents
    KillApp "PDFGen.exe" ' Cargar lista de procesos activos (KillApp)
    DoEvents
'        Enviar_Mail_CDO "julio.gonzalez@ixitec.net", "HAY MAS DE 10 PROCESOS PETADOS", ""
'        MsgBox "HAY MAS DE 10 PROCESOS PETADOS. PARADA TECNICA....", vbCritical, App.Title
'    End If
    ' Verificamos si hay informes pendientes y si es asi no recargamos la lista
    cargar_lista
    If lstActive.ListCount < 15 Then
        imprimir
    End If
End Sub

Private Sub imprimir()
    On Error Resume Next
    Dim i As Integer
    Dim IMPRESORA As Integer
    tot = "Total : " & Lista.ListItems.Count
    For i = 1 To Lista.ListItems.Count
        If Lista.ListItems(i).SubItems(4) = 1 Then
            Exit For
        End If
        If Lista.ListItems(i).SubItems(4) = 0 Then
          Lista.ListItems(i).SubItems(4) = 1
          DoEvents
          Shell App.Path & "\PDFGen.exe" & " " & Lista.ListItems(i).SubItems(8) & " " & database
          TrayIcon1.ShowBalloonTip 3, "Generando Informe... " & Lista.ListItems(i).SubItems(8), "Solicitando generación : " & Lista.ListItems(i).SubItems(8), 0
          Exit Sub
        End If
    Next
End Sub

Private Sub TrayIcon1_DblClick()
    If (Minimized) Then Command1_Click
End Sub
