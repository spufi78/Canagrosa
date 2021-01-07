VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversión BBDD Metrología"
   ClientHeight    =   10230
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   9765
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6044.223
   ScaleMode       =   0  'User
   ScaleWidth      =   9168.806
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.WebBrowser WebBrowser1 
      Height          =   30
      Left            =   9360
      TabIndex        =   17
      Top             =   180
      Width           =   30
      _Version        =   851970
      _ExtentX        =   53
      _ExtentY        =   53
      _StockProps     =   173
      BackColor       =   -2147483643
   End
   Begin VB.OptionButton opTipo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MANOMETRO (CARPETA \\SERVIDOR\CANAGROSA\METROLOGIA\MANOMETRO)"
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
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   360
      Width           =   7665
   End
   Begin VB.OptionButton opTipo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PIE DE REY (CARPETA \\SERVIDOR\CANAGROSA\METROLOGIA\PIEDEREY)"
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
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   90
      Value           =   -1  'True
      Width           =   7665
   End
   Begin VB.CheckBox chklog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Activar Log"
      Height          =   240
      Left            =   45
      TabIndex        =   14
      Top             =   9225
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.TextBox txtLinea 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9090
      TabIndex        =   12
      Text            =   "5"
      Top             =   135
      Width           =   645
   End
   Begin VB.TextBox txtlog 
      Appearance      =   0  'Flat
      Height          =   8115
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1035
      Width           =   9690
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   8460
      Picture         =   "frmLogin.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9225
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Convertir"
      Height          =   915
      Left            =   7155
      Picture         =   "frmLogin.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9225
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3840
      Left            =   990
      TabIndex        =   2
      Top             =   4230
      Visible         =   0   'False
      Width           =   7395
      Begin VB.CheckBox chkPrueba 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Iniciar en modo Prueba"
         Height          =   240
         Left            =   1845
         TabIndex        =   6
         Top             =   2385
         Width           =   2085
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1830
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1995
         Width           =   1965
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1830
         TabIndex        =   0
         Top             =   1605
         Width           =   1965
      End
      Begin VB.Label lblreg 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   450
         TabIndex        =   5
         Top             =   3690
         Width           =   3855
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   450
         TabIndex        =   4
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H80000005&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   450
         TabIndex        =   3
         Top             =   1620
         Width           =   1260
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fila Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8055
      TabIndex        =   13
      Top             =   225
      Width           =   1185
   End
   Begin VB.Label lblplanta 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   585
      TabIndex        =   11
      Top             =   9540
      Width           =   1680
   End
   Begin VB.Label lblregistro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   585
      TabIndex        =   10
      Top             =   9855
      Width           =   1680
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PIE_DE_REY = "PIEDEREY"
Const MANOMETRO = "MANOMETRO"
Private Sub cmdcancel_Click()
    End
End Sub

Private Sub cmdok_Click() 'comprobar si la contraseña es correcta
    Me.MousePointer = 11
    Set USUARIO = New ClsUsuario
    If CrearConexionGlobal(txtdatos(0), ReadINI(App.Path + "\config.ini", "SERVER", "BD_USUARIO"), ReadINI(App.Path + "\config.ini", "SERVER", "BD_PASS"), chkPrueba.Value) = True Then
        If opTipo(0).Value = True Then ' PIE_DE_REY
            cargar_lista_xlsx (PIE_DE_REY)
        End If
        If opTipo(1).Value = True Then ' MANOMETRO
            cargar_lista_xlsx (MANOMETRO)
        End If
    Else
        Me.MousePointer = 0
        MsgBox "No se pudo conectar con la base de datos", vbCritical, Err.Description
    End If
    Me.MousePointer = 0
    MsgBox "OK"
End Sub
Private Sub cargar_lista_xlsx(carpeta As String)
    Dim fso, fsoFolder, fsoFile, strPath, strName, fsoSubFolder, i, a
    i = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFolder = fso.GetFolder(App.Path & "\" & carpeta)
    
    For Each fsoFile In fsoFolder.Files
        If LCase(Right(fsoFile.Name, 5)) = ".xlsx" Or LCase(Right(fsoFile.Name, 5)) = ".xlsm" Then
            txtlog = txtlog & "PROCESANDO : " & fsoFile.Name & vbNewLine
            If carpeta = PIE_DE_REY Then
                excel_pie_de_rey fsoFile.Name
            End If
            If carpeta = MANOMETRO Then
                excel_manometro fsoFile.Name
            End If
        End If
    Next
End Sub

Private Sub Form_Load()
    On Error Resume Next
    log (Me.Name)
    txtdatos(0) = ReadINI(App.Path + "\config.ini", "usuario", "usuario")
    txtdatos(1) = ReadINI(App.Path + "\config.ini", "usuario", "pass")
    ' Para Pruebas
'    registrar_componentes (Me.Hwnd)
    If txtdatos(0) <> "" Then
'        cmdok_Click
    End If
End Sub

Private Sub opTipo_Click(Index As Integer)
    Select Case Index
    Case 0
        txtLinea = "5"
    Case 1
        txtLinea = "8"
    End Select
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    If Index <> 2 Then
        txtdatos(Index).SelStart = 0
        txtdatos(Index).SelLength = Len(txtdatos(Index))
    End If
End Sub

Private Sub txtdatos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40
        SendKeys "{Tab}", True
     Case 38
        SendKeys "+{Tab}", True
     Case 27
        cmdcancel_Click
    End Select
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
End Sub

 Private Function Aleatorio(Minimo As Long, maximo As Long) As Long
     Randomize ' inicializar la semilla
     Aleatorio = CLng((Minimo - maximo) * Rnd + maximo)
 End Function
Public Sub excel_pie_de_rey(hoja As String)
    Dim XLA As Excel.Application
    Dim XLW As Excel.Workbook
    Dim XLS As Excel.Worksheet
   On Error GoTo excel_pie_de_rey_Error

    Set XLA = New Excel.Application
    Set XLW = XLA.Workbooks.Open(App.Path & "\" & PIE_DE_REY & "\" & hoja)
    Dim activa As Integer
    activa = 1
    Dim j As Integer
    For j = 1 To XLW.Worksheets.Count
        txtlog = j & " -> " & XLW.Worksheets(j).Name & vbNewLine & txtlog
        If UCase(XLW.Worksheets(j).Name) = "DATOS PIE DE REY" Then
            activa = j
        End If
    Next
    Set XLS = XLW.Worksheets(activa)
    Dim rsEquipo As ADODB.Recordset
    Dim i As Integer
    Dim c As String
    Dim filaCodigo As Integer
    filaCodigo = 2
    i = txtLinea.Text
    While XLS.Cells(i, filaCodigo) <> "" And XLS.Cells(i, filaCodigo) <> "0"
        If chklog.Value = Checked Then
            txtlog = "Procesando fila : " & i & " Codigo : " & XLS.Cells(i, filaCodigo) & vbNewLine & txtlog
        End If
        c = "select * from geslab_canagrosa.equipos where convert(numero_equipo_cliente using latin1) = '" & XLS.Cells(i, filaCodigo) & "' and cliente_id <> 0 "
        Set rsEquipo = datos_bd(c)
        If rsEquipo.RecordCount = 0 Then
            If chklog.Value = Checked Then
                txtlog = "No existe el equipo : " & XLS.Cells(i, filaCodigo) & vbNewLine & txtlog
            End If
        Else
            Dim oPR As New clsPie_de_rey
            With oPR
                .setEQUIPO_ID = rsEquipo("ID_EQUIPO")
                .setNOMINAL_ANCHO_COMBINADO_L4 = ""
                If XLS.Cells(i, 11) <> "" Then
                    If IsNumeric(XLS.Cells(i, 11)) Then
                        .setNOMINAL_ANCHO_COMBINADO_L4 = XLS.Cells(i, 11)
                    End If
                End If
                .setTOLERANCIA_ANCHO_COMBINADO_L4_MM = ""
                If XLS.Cells(i, 12) <> "" Then
                    If IsNumeric(XLS.Cells(i, 12)) Then
                        .setTOLERANCIA_ANCHO_COMBINADO_L4_MM = Replace(XLS.Cells(i, 12), ",", ".")
                    End If
                End If
                .setFRASE_CRITERIO_ACEPTACION_ESP = XLS.Cells(i, 15)
                .setFRASE_CRITERIO_ACEPTACION_ENG = XLS.Cells(i, 16)
                .setLONGITUD_BOCASH = ""
                If XLS.Cells(i, 17) <> "" Then
                    If IsNumeric(XLS.Cells(i, 17)) Then
                        .setLONGITUD_BOCASH = Replace(XLS.Cells(i, 17), ",", ".")
                    End If
                End If
                .setLONGITUD_BOCAS_H = ""
                If XLS.Cells(i, 18) <> "" Then
                    If IsNumeric(XLS.Cells(i, 18)) Then
                        .setLONGITUD_BOCAS_H = Replace(XLS.Cells(i, 18), ",", ".")
                    End If
                End If
                
                .setPUNTO_DE_CALIBRACION_N1_PALPADOR_DE_EXTERIORES = ""
                .setPUNTO_DE_CALIBRACION_N2_PALPADOR_DE_EXTERIORES = ""
                .setPUNTO_DE_CALIBRACION_N3_PALPADOR_DE_EXTERIORES = ""
                .setPUNTO_DE_CALIBRACION_N4_PALPADOR_DE_EXTERIORES = ""
                .setPUNTO_DE_CALIBRACION_N5_PALPADOR_DE_EXTERIORES = ""
                .setPUNTO_DE_CALIBRACION_N6_PALPADOR_DE_EXTERIORES = ""
                .setPUNTO_DE_CALIBRACION_N7_PALPADOR_DE_EXTERIORES = ""
                .setPUNTO_DE_CALIBRACION_N8_PALPADOR_DE_EXTERIORES = ""
                .setPUNTO_DE_CALIBRACION_N9_PALPADOR_DE_EXTERIORES = ""
                .setPUNTO_DE_CALIBRACION_N10_PALPADOR_DE_EXTERIORES = ""
                
                Dim h As Integer
                Dim v As String
                For h = 19 To 28
                    v = XLS.Cells(i, h)
                    If v <> "" Then
                        If IsNumeric(v) Then
                            Select Case h - 18
                            Case 1
                                .setPUNTO_DE_CALIBRACION_N1_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            Case 2
                                .setPUNTO_DE_CALIBRACION_N2_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            Case 3
                                .setPUNTO_DE_CALIBRACION_N3_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            Case 4
                                .setPUNTO_DE_CALIBRACION_N4_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            Case 5
                                .setPUNTO_DE_CALIBRACION_N5_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            Case 6
                                .setPUNTO_DE_CALIBRACION_N6_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            Case 7
                                .setPUNTO_DE_CALIBRACION_N7_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            Case 8
                                .setPUNTO_DE_CALIBRACION_N8_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            Case 9
                                .setPUNTO_DE_CALIBRACION_N9_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            Case 10
                                .setPUNTO_DE_CALIBRACION_N10_PALPADOR_DE_EXTERIORES = Replace(v, ",", ".")
                            End Select
                        End If
                    End If
                Next
                
               .setPUNTO_DE_CALIBRACION_N1_PALPADOR_DE_INTERIORES = ""
               .setPUNTO_DE_CALIBRACION_N2_PALPADOR_DE_INTERIORES = ""
               .setPUNTO_DE_CALIBRACION_N3_PALPADOR_DE_INTERIORES = ""

                For h = 29 To 31
                    v = XLS.Cells(i, h)
                    If v <> "" Then
                        If IsNumeric(v) Then
                            Select Case h - 28
                            Case 1
                                .setPUNTO_DE_CALIBRACION_N1_PALPADOR_DE_INTERIORES = Replace(v, ",", ".")
                            Case 2
                                .setPUNTO_DE_CALIBRACION_N2_PALPADOR_DE_INTERIORES = Replace(v, ",", ".")
                            Case 3
                                .setPUNTO_DE_CALIBRACION_N3_PALPADOR_DE_INTERIORES = Replace(v, ",", ".")
                            End Select
                        End If
                    End If
                Next
               .setPUNTO_DE_CALIBRACION_N1_SONDA_DE_PROFUNDIDAD = ""
               .setPUNTO_DE_CALIBRACION_N2_SONDA_DE_PROFUNDIDAD = ""
               .setPUNTO_DE_CALIBRACION_N3_SONDA_DE_PROFUNDIDAD = ""

                For h = 32 To 34
                    v = XLS.Cells(i, h)
                    If v <> "" Then
                        If IsNumeric(v) Then
                            Select Case h - 28
                            Case 1
                                .setPUNTO_DE_CALIBRACION_N1_SONDA_DE_PROFUNDIDAD = Replace(v, ",", ".")
                            Case 2
                                .setPUNTO_DE_CALIBRACION_N2_SONDA_DE_PROFUNDIDAD = Replace(v, ",", ".")
                            Case 3
                                .setPUNTO_DE_CALIBRACION_N3_SONDA_DE_PROFUNDIDAD = Replace(v, ",", ".")
                            End Select
                        End If
                    End If
                Next
                
                .setCRITERIO_C = ""
                If XLS.Cells(i, 42) <> "" Then
                    If IsNumeric(XLS.Cells(i, 42)) Then
                        .setCRITERIO_C = Replace(XLS.Cells(i, 42), ",", ".")
                    End If
                End If
                .setCRITERIO_UEXP = ""
                If XLS.Cells(i, 43) <> "" Then
                    If IsNumeric(XLS.Cells(i, 43)) Then
                        .setCRITERIO_UEXP = Replace(XLS.Cells(i, 43), ",", ".")
                    End If
                End If
                .setCRITERIO_C_UEXP = ""
                If XLS.Cells(i, 44) <> "" Then
                    If IsNumeric(XLS.Cells(i, 44)) Then
                        .setCRITERIO_C_UEXP = Replace(XLS.Cells(i, 44), ",", ".")
                    End If
                End If

                .Insertar
            End With
        End If
        
        i = i + 1
        lblregistro = i
        If chklog.Value = Checked Then
'            txtlog = ""
            DoEvents
        Else
            If i Mod 100 = 0 Then
                DoEvents
            End If
        End If
    Wend
    DoEvents
    XLW.Close
    XLA.Quit
    txtlog = txtlog & "FINALIZADA HOJA :" & hoja & vbNewLine
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

excel_pie_de_rey_Error:
    XLW.Close
    XLA.Quit
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure excel_pie_de_rey of Formulario frmLogin"
End Sub

Public Sub excel_manometro(hoja As String)
    Dim XLA As Excel.Application
    Dim XLW As Excel.Workbook
    Dim XLS As Excel.Worksheet
   On Error GoTo excel_pie_de_rey_Error

    Set XLA = New Excel.Application
    Set XLW = XLA.Workbooks.Open(App.Path & "\" & MANOMETRO & "\" & hoja)
    Dim activa As Integer
    activa = 1
    Dim j As Integer
    For j = 1 To XLW.Worksheets.Count
        txtlog = j & " -> " & XLW.Worksheets(j).Name & vbNewLine & txtlog
        If UCase(XLW.Worksheets(j).Name) = "DATOS" Then
            activa = j
        End If
    Next
    Set XLS = XLW.Worksheets(activa)
    Dim rsEquipo As ADODB.Recordset
    Dim i As Integer
    Dim c As String
    Dim filaCodigo As Integer
    filaCodigo = 3
    i = txtLinea.Text
    While XLS.Cells(i, filaCodigo) <> "" And XLS.Cells(i, filaCodigo) <> "0"
        If chklog.Value = Checked Then
            txtlog = "Procesando fila : " & i & " Codigo : " & XLS.Cells(i, filaCodigo) & vbNewLine & txtlog
        End If
        c = "select * from geslab_canagrosa.equipos where convert(numero_equipo_cliente using latin1) = '" & XLS.Cells(i, filaCodigo) & "' and cliente_id <> 0 "
        Set rsEquipo = datos_bd(c)
        If rsEquipo.RecordCount = 0 Then
            If chklog.Value = Checked Then
                txtlog = "No existe el equipo : " & XLS.Cells(i, filaCodigo) & vbNewLine & txtlog
            End If
        Else
            ' Si el tipo de equipo no es MANOMETRO lo modificacion
            If rsEquipo("TIPO_EQUIPO_ID") <> 9 Then
                If chklog.Value = Checked Then
                    txtlog = "El equipo no es de tipo manometro, modificando : " & XLS.Cells(i, filaCodigo) & vbNewLine & txtlog
                    execute_bd "UPDATE geslab_canagrosa.equipos set TIPO_EQUIPO_ID = 9 where ID_EQUIPO = " & rsEquipo("ID_EQUIPO")
                End If
            End If
            Dim oManometro As New clsManometro_vacuometro
            With oManometro
                .setEQUIPO_ID = rsEquipo("ID_EQUIPO")
                .setFLUIDO_DE_CAL = XLS.Cells(i, 20)
                .setFLUIDO_CAL_INGLES = XLS.Cells(i, 21)
                .setCRITERIO_ACEPTACION = XLS.Cells(i, 30)
                .setACCEPTANCE_CRITERIA = XLS.Cells(i, 34)
                ' PNT F
                .setPNT_F_ASOCIADO = ""
                Select Case XLS.Cells(i, 63)
                    Case "001"
                        .setPNT_F_ASOCIADO = 2709
                    Case "002"
                        .setPNT_F_ASOCIADO = 2710
                    Case "003"
                        .setPNT_F_ASOCIADO = 2711
                    Case "004"
                        .setPNT_F_ASOCIADO = 2712
                    Case "005"
                        .setPNT_F_ASOCIADO = 2713
                    Case "006"
                        .setPNT_F_ASOCIADO = 2733
                    Case "007"
                        .setPNT_F_ASOCIADO = 2734
                    Case "008"
                        .setPNT_F_ASOCIADO = 2735
                    Case "010"
                        .setPNT_F_ASOCIADO = 2736
                    Case "011"
                        .setPNT_F_ASOCIADO = 2737
                    Case "012"
                        .setPNT_F_ASOCIADO = 2738
                    Case "013"
                        .setPNT_F_ASOCIADO = 2739
                    Case "014"
                        .setPNT_F_ASOCIADO = 2740
                    Case "015"
                        .setPNT_F_ASOCIADO = 2741
                    Case "016"
                        .setPNT_F_ASOCIADO = 2742
                    Case "017"
                        .setPNT_F_ASOCIADO = 2743
                    Case "018"
                        .setPNT_F_ASOCIADO = 2744
                    Case "019"
                        .setPNT_F_ASOCIADO = 2745
                    Case "020"
                        .setPNT_F_ASOCIADO = 2746
                    Case "025"
                        .setPNT_F_ASOCIADO = 2748
                    Case "027"
                        .setPNT_F_ASOCIADO = 2788
                    Case "028"
                        .setPNT_F_ASOCIADO = 2789
                    Case "030"
                        .setPNT_F_ASOCIADO = 2791
                    Case "031"
                        .setPNT_F_ASOCIADO = 2792
                    Case "032"
                        .setPNT_F_ASOCIADO = 2793
                    Case "033"
                        .setPNT_F_ASOCIADO = 2794
                    Case "034"
                        .setPNT_F_ASOCIADO = 2795
                    Case "035"
                        .setPNT_F_ASOCIADO = 2796
                    Case "036"
                        .setPNT_F_ASOCIADO = 2797
                    Case "041"
                        .setPNT_F_ASOCIADO = 2286
                    Case "048"
                        .setPNT_F_ASOCIADO = 2396
                    Case "049"
                        .setPNT_F_ASOCIADO = 2397
                End Select
                .setRANGO_UTIL_CAL = ""
                If XLS.Cells(i, 74) <> "" Then
                    If IsNumeric(XLS.Cells(i, 74)) Then
                        .setRANGO_UTIL_CAL = Replace(XLS.Cells(i, 74), ",", ".")
                    End If
                End If

                .setUNIDAD_TOLERANCIA = XLS.Cells(i, 85)
                .setEVALUACION_TOLERANCIA = XLS.Cells(i, 86)
                .setUNIDAD_AJUSTE = XLS.Cells(i, 97)
                .setEVALUACION_AJUSTE = XLS.Cells(i, 98)
                .setTRAZMET = XLS.Cells(i, 99)
                .setAUXILIAR2 = XLS.Cells(i, 100)
                .setAUXILIAR3 = XLS.Cells(i, 101)
                .setINTERVALO_CALIBRACION_MESES = XLS.Cells(i, 102)
                .setI_USO = XLS.Cells(i, 103)
                .setNOMBRE_BANCO = XLS.Cells(i, 104)
                .setIDENTIFICACION_BANCO = XLS.Cells(i, 105)
                .setCLASE_DE_INFORME = XLS.Cells(i, 127)
                                
                .setPUNTO1 = ""
                .setPUNTO2 = ""
                .setPUNTO3 = ""
                .setPUNTO4 = ""
                .setPUNTO5 = ""
                .setPUNTO6 = ""
                .setPUNTO7 = ""
                .setPUNTO8 = ""
                .setPUNTO9 = ""
                .setPUNTO10 = ""

                Dim h As Integer
                Dim v As String
                For h = 38 To 47
                    v = XLS.Cells(i, h)
                    If v <> "" Then
                        If IsNumeric(v) Then
                            Select Case h - 37
                            Case 1
                                .setPUNTO1 = Replace(v, ",", ".")
                            Case 2
                                .setPUNTO2 = Replace(v, ",", ".")
                            Case 3
                                .setPUNTO3 = Replace(v, ",", ".")
                            Case 4
                                .setPUNTO4 = Replace(v, ",", ".")
                            Case 5
                                .setPUNTO5 = Replace(v, ",", ".")
                            Case 6
                                .setPUNTO6 = Replace(v, ",", ".")
                            Case 7
                                .setPUNTO7 = Replace(v, ",", ".")
                            Case 8
                                .setPUNTO8 = Replace(v, ",", ".")
                            Case 9
                                .setPUNTO9 = Replace(v, ",", ".")
                            Case 10
                                .setPUNTO10 = Replace(v, ",", ".")
                            End Select
                        End If
                    End If
                Next
                
                .setCMC_PUNTO1 = ""
                .setCMC_PUNTO2 = ""
                .setCMC_PUNTO3 = ""
                .setCMC_PUNTO4 = ""
                .setCMC_PUNTO5 = ""
                .setCMC_PUNTO6 = ""
                .setCMC_PUNTO7 = ""
                .setCMC_PUNTO8 = ""
                .setCMC_PUNTO9 = ""
                .setCMC_PUNTO10 = ""

                For h = 64 To 73
                    v = XLS.Cells(i, h)
                    If v <> "" Then
                        If IsNumeric(v) Then
                            Select Case h - 63
                            Case 1
                                .setCMC_PUNTO1 = Replace(v, ",", ".")
                            Case 2
                                .setCMC_PUNTO2 = Replace(v, ",", ".")
                            Case 3
                                .setCMC_PUNTO3 = Replace(v, ",", ".")
                            Case 4
                                .setCMC_PUNTO4 = Replace(v, ",", ".")
                            Case 5
                                .setCMC_PUNTO5 = Replace(v, ",", ".")
                            Case 6
                                .setCMC_PUNTO6 = Replace(v, ",", ".")
                            Case 7
                                .setCMC_PUNTO7 = Replace(v, ",", ".")
                            Case 8
                                .setCMC_PUNTO8 = Replace(v, ",", ".")
                            Case 9
                                .setCMC_PUNTO9 = Replace(v, ",", ".")
                            Case 10
                                .setCMC_PUNTO10 = Replace(v, ",", ".")
                            End Select
                        End If
                    End If
                Next
                
                .setTOLERANCIA_PUNTO1 = ""
                .setTOLERANCIA_PUNTO2 = ""
                .setTOLERANCIA_PUNTO3 = ""
                .setTOLERANCIA_PUNTO4 = ""
                .setTOLERANCIA_PUNTO5 = ""
                .setTOLERANCIA_PUNTO6 = ""
                .setTOLERANCIA_PUNTO7 = ""
                .setTOLERANCIA_PUNTO8 = ""
                .setTOLERANCIA_PUNTO9 = ""
                .setTOLERANCIA_PUNTO10 = ""

                For h = 75 To 84
                    v = XLS.Cells(i, h)
                    If v <> "" Then
                        If IsNumeric(v) Then
                            Select Case h - 74
                            Case 1
                                .setTOLERANCIA_PUNTO1 = Replace(v, ",", ".")
                            Case 2
                                .setTOLERANCIA_PUNTO2 = Replace(v, ",", ".")
                            Case 3
                                .setTOLERANCIA_PUNTO3 = Replace(v, ",", ".")
                            Case 4
                                .setTOLERANCIA_PUNTO4 = Replace(v, ",", ".")
                            Case 5
                                .setTOLERANCIA_PUNTO5 = Replace(v, ",", ".")
                            Case 6
                                .setTOLERANCIA_PUNTO6 = Replace(v, ",", ".")
                            Case 7
                                .setTOLERANCIA_PUNTO7 = Replace(v, ",", ".")
                            Case 8
                                .setTOLERANCIA_PUNTO8 = Replace(v, ",", ".")
                            Case 9
                                .setTOLERANCIA_PUNTO9 = Replace(v, ",", ".")
                            Case 10
                                .setTOLERANCIA_PUNTO10 = Replace(v, ",", ".")
                            End Select
                        End If
                    End If
                Next
                
                .setCRITERIO_DE_AJUSTE_PUNTO1 = ""
                .setCRITERIO_DE_AJUSTE_PUNTO2 = ""
                .setCRITERIO_DE_AJUSTE_PUNTO3 = ""
                .setCRITERIO_DE_AJUSTE_PUNTO4 = ""
                .setCRITERIO_DE_AJUSTE_PUNTO5 = ""
                .setCRITERIO_DE_AJUSTE_PUNTO6 = ""
                .setCRITERIO_DE_AJUSTE_PUNTO7 = ""
                .setCRITERIO_DE_AJUSTE_PUNTO8 = ""
                .setCRITERIO_DE_AJUSTE_PUNTO9 = ""
                .setCRITERIO_DE_AJUSTE_PUNTO10 = ""

                For h = 87 To 96
                    v = XLS.Cells(i, h)
                    If v <> "" Then
                        If IsNumeric(v) Then
                            Select Case h - 86
                            Case 1
                                .setCRITERIO_DE_AJUSTE_PUNTO1 = Replace(v, ",", ".")
                            Case 2
                                .setCRITERIO_DE_AJUSTE_PUNTO2 = Replace(v, ",", ".")
                            Case 3
                                .setCRITERIO_DE_AJUSTE_PUNTO3 = Replace(v, ",", ".")
                            Case 4
                                .setCRITERIO_DE_AJUSTE_PUNTO4 = Replace(v, ",", ".")
                            Case 5
                                .setCRITERIO_DE_AJUSTE_PUNTO5 = Replace(v, ",", ".")
                            Case 6
                                .setCRITERIO_DE_AJUSTE_PUNTO6 = Replace(v, ",", ".")
                            Case 7
                                .setCRITERIO_DE_AJUSTE_PUNTO7 = Replace(v, ",", ".")
                            Case 8
                                .setCRITERIO_DE_AJUSTE_PUNTO8 = Replace(v, ",", ".")
                            Case 9
                                .setCRITERIO_DE_AJUSTE_PUNTO9 = Replace(v, ",", ".")
                            Case 10
                                .setCRITERIO_DE_AJUSTE_PUNTO10 = Replace(v, ",", ".")
                            End Select
                        End If
                    End If
                Next
                
                .setCMC_TOL_PUNTO_1 = ""
                .setCMC_TOL_PUNTO_2 = ""
                .setCMC_TOL_PUNTO_3 = ""
                .setCMC_TOL_PUNTO_4 = ""
                .setCMC_TOL_PUNTO_5 = ""
                .setCMC_TOL_PUNTO_6 = ""
                .setCMC_TOL_PUNTO_7 = ""
                .setCMC_TOL_PUNTO_8 = ""
                .setCMC_TOL_PUNTO_9 = ""
                .setCMC_TOL_PUNTO_10 = ""
                
                For h = 106 To 115
                    If IsNumeric(XLS.Cells(i, h)) Then
                        v = XLS.Cells(i, h)
                        If v <> "" Then
                            If IsNumeric(v) Then
                                Select Case h - 105
                                Case 1
                                    .setCMC_TOL_PUNTO_1 = Replace(v, ",", ".")
                                Case 2
                                    .setCMC_TOL_PUNTO_2 = Replace(v, ",", ".")
                                Case 3
                                    .setCMC_TOL_PUNTO_3 = Replace(v, ",", ".")
                                Case 4
                                    .setCMC_TOL_PUNTO_4 = Replace(v, ",", ".")
                                Case 5
                                    .setCMC_TOL_PUNTO_5 = Replace(v, ",", ".")
                                Case 6
                                    .setCMC_TOL_PUNTO_6 = Replace(v, ",", ".")
                                Case 7
                                    .setCMC_TOL_PUNTO_7 = Replace(v, ",", ".")
                                Case 8
                                    .setCMC_TOL_PUNTO_8 = Replace(v, ",", ".")
                                Case 9
                                    .setCMC_TOL_PUNTO_9 = Replace(v, ",", ".")
                                Case 10
                                    .setCMC_TOL_PUNTO_10 = Replace(v, ",", ".")
                                End Select
                            End If
                        End If
                    End If
                Next
                
                .setTRANSD_PUNTO1 = ""
                .setTRANSD_PUNTO2 = ""
                .setTRANSD_PUNTO3 = ""
                .setTRANSD_PUNTO4 = ""
                .setTRANSD_PUNTO5 = ""
                .setTRANSD_PUNTO6 = ""
                .setTRANSD_PUNTO7 = ""
                .setTRANSD_PUNTO8 = ""
                .setTRANSD_PUNTO9 = ""
                .setTRANSD_PUNTO10 = ""
                
                For h = 116 To 125
                    v = XLS.Cells(i, h)
                    If v <> "" Then
                        If IsNumeric(v) Then
                            Select Case h - 115
                            Case 1
                                .setTRANSD_PUNTO1 = Replace(v, ",", ".")
                            Case 2
                                .setTRANSD_PUNTO2 = Replace(v, ",", ".")
                            Case 3
                                .setTRANSD_PUNTO3 = Replace(v, ",", ".")
                            Case 4
                                .setTRANSD_PUNTO4 = Replace(v, ",", ".")
                            Case 5
                                .setTRANSD_PUNTO5 = Replace(v, ",", ".")
                            Case 6
                                .setTRANSD_PUNTO6 = Replace(v, ",", ".")
                            Case 7
                                .setTRANSD_PUNTO7 = Replace(v, ",", ".")
                            Case 8
                                .setTRANSD_PUNTO8 = Replace(v, ",", ".")
                            Case 9
                                .setTRANSD_PUNTO9 = Replace(v, ",", ".")
                            Case 10
                                .setTRANSD_PUNTO10 = Replace(v, ",", ".")
                            End Select
                        End If
                    End If
                Next

                .setTRANS1_CERT = ""
                .setTRANS2_CERT = ""
                .setTRANS3_CERT = ""
                For h = 128 To 130
                    v = XLS.Cells(i, h)
                    If v <> "" Then
                        If IsNumeric(v) Then
                            Select Case h - 127
                            Case 1
                                .setTRANS1_CERT = Replace(v, ",", ".")
                            Case 2
                                .setTRANS2_CERT = Replace(v, ",", ".")
                            Case 3
                                .setTRANS3_CERT = Replace(v, ",", ".")
                            End Select
                        End If
                    End If
                Next

                .Insertar
            End With
        End If
        
        i = i + 1
        lblregistro = i
        If chklog.Value = Checked Then
'            txtlog = ""
            DoEvents
        Else
            If i Mod 100 = 0 Then
                DoEvents
            End If
        End If
    Wend
    DoEvents
    XLW.Close 0
    XLA.Quit
    txtlog = txtlog & "FINALIZADA HOJA :" & hoja & vbNewLine
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

excel_pie_de_rey_Error:
    XLW.Close 0
    XLA.Quit
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure excel_pie_de_rey of Formulario frmLogin"
End Sub

