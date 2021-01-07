VERSION 5.00
Object = "{F4375239-2DAA-489A-9DCE-662FC9185BD6}#1.99#0"; "BarcodeWiz.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEtiquetas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Etiquetas"
   ClientHeight    =   6975
   ClientLeft      =   4590
   ClientTop       =   2880
   ClientWidth     =   6885
   Icon            =   "frmEtiquetas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6885
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1635
      TabIndex        =   12
      Top             =   5505
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1635
      TabIndex        =   10
      Top             =   5190
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1635
      TabIndex        =   8
      Top             =   4875
      Width           =   990
   End
   Begin VB.OptionButton optSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pequeña"
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Value           =   -1  'True
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mediana"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   5220
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grande"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   5550
      Width           =   972
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   915
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6030
      Width           =   1275
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   3630
      TabIndex        =   3
      Top             =   4080
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtDatos(0)"
      BuddyDispid     =   196613
      BuddyIndex      =   0
      OrigLeft        =   2820
      OrigTop         =   960
      OrigRight       =   3060
      OrigBottom      =   1395
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2955
      TabIndex        =   2
      Text            =   "1"
      Top             =   4080
      Width           =   675
   End
   Begin MSComctlLib.ListView lista 
      Height          =   1785
      Left            =   3990
      TabIndex        =   15
      Top             =   4140
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3149
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
   Begin BARCODEWIZLibCtl.BarCodeWiz Barcode 
      Height          =   810
      Left            =   2460
      TabIndex        =   17
      Top             =   1050
      Width           =   1545
      m_scaleNumerator=   1
      m_scaleDenominator=   1
      _cx             =   2735
      _cy             =   1431
      AutoSize        =   -1  'True
      BackColor       =   16777215
      BackStyle       =   1
      Barcode         =   "1234"
      BarcodeHeight   =   600
      BeginProperty BarcodeTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BarcodeTextPosition=   0
      BearerBars      =   0   'False
      Border          =   0
      BottomText      =   "Abajo"
      BeginProperty BottomTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BottomTextAlignment=   2
      Enabled         =   -1  'True
      ForeColor       =   0
      NarrowBarWidth  =   35
      OptionalCheckChar=   0
      Orientation     =   0
      QuietZone       =   1
      ScaleMode       =   0
      StretchBarcodeText=   0   'False
      Symbology       =   6
      TopText         =   "Arriba"
      TopTextAlignment=   2
      BeginProperty TopTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WideToNarrowRatio=   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Previsualización de etiqueta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   60
      Width           =   6645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tamaño"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   1770
      TabIndex        =   14
      Top             =   4605
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "px"
      Height          =   195
      Index           =   4
      Left            =   2700
      TabIndex        =   13
      Top             =   5550
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "px"
      Height          =   195
      Index           =   2
      Left            =   2700
      TabIndex        =   11
      Top             =   5235
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "px"
      Height          =   195
      Index           =   0
      Left            =   2700
      TabIndex        =   9
      Top             =   4920
      Width           =   165
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Etiquetas a imprimir por muestra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   4
      Top             =   4155
      Width           =   2715
   End
End
Attribute VB_Name = "frmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim muestra_proceso As Long
Dim linea_lista As Integer
Private Sub Barcode_Click()
    Barcode.ShowProperties
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo errores
    ' Imprimimos las etiquetas
    Dim imp_ant As String
'    imp_ant = Impresora_Predeterminada
    Dim oParametro As New clsParametros
    If Not oParametro.Carga(parametros.IMPRESORA_ETIQUETAS_PEQUENA, USUARIO.getUSO) Then
        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
        Exit Sub
    End If
    log ("Comienzo impresion de etiquetas")
    Dim encontrada As Boolean
    encontrada = False
    Dim impresora_encontrada As String
    impresora_encontrada = ""
    For Each prnPrinter In Printers
        ' Quitamos del nombre de la impresora el #00X para las impresoras TSN del remoto
        Dim impresoraArray() As String
        Dim impresoraNombre As String
        impresoraArray = Split(UCase(prnPrinter.DeviceName), "#")
        impresoraNombre = Trim(impresoraArray(0))
        
'        If UCase(prnPrinter.DeviceName) = Replace(UCase(oParametro.getVALOR), "/", "\") Then
        If impresoraNombre = Replace(UCase(oParametro.getVALOR), "/", "\") Then
            encontrada = True
            Set Printer = prnPrinter
            impresora_encontrada = prnPrinter.DeviceName
            Exit For
        End If
    Next
    If Not encontrada Then
        MsgBox "No se encuentra la impresora de etiquetas.", vbCritical, App.Title
        Exit Sub
    End If
    Dim i As Integer
    Dim j As Integer
'    Establecer_Impresora impresora_encontrada
    For j = 1 To lista.ListItems.Count
        For i = 1 To txtDatos(0)
            linea_lista = j
            If optSize(0).value = True Then
                optSize_Click (0)
            End If
            If optSize(1).value = True Then
                optSize_Click (1)
            End If
            If optSize(2).value = True Then
                optSize_Click (2)
            End If
            Printer.PaintPicture Barcode.Picture, 0, 0
            Printer.NewPage
        Next
    Next
    Printer.EndDoc
    log ("Final impresion de etiquetas")
'    Establecer_Impresora imp_ant
    Exit Sub
errores:
    MsgBox "Error al imprimir las etiquetas.", vbCritical, Err.Description
End Sub

Private Sub cmdSalir_Click()
    ReDim etiquetas(1)
    etiquetas(1) = 0
    muestra_proceso = 0
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27
        cmdSalir_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_medidas
    If etiquetas(1) = 0 Then
        MsgBox "No se ha informado el valor.", vbCritical, App.Title
    Else
        With lista.ColumnHeaders
            .Add , , "ID", 1, lvwColumnLeft
            .Add , , "Muestra", 2500, lvwColumnCenter
        End With
        cargar_lista
'        muestra_proceso = etiquetas(1)
        linea_lista = 1
        optSize_Click 0
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        linea_lista = lista.selectedItem.Index
'        muestra_proceso = lista.ListItems(lista.SelectedItem.Index).Text
        If optSize(0).value = True Then
            optSize_Click 0
        ElseIf optSize(1).value = True Then
            optSize_Click 1
        Else
            optSize_Click 2
        End If
    End If
End Sub

Private Sub optSize_Click(Index As Integer)
    Dim oMuestra As New clsMuestra
'    oMuestra.CargaMuestra (muestra_proceso)
    oMuestra.CargaMuestra lista.ListItems(linea_lista).Text
    With Barcode
        .BottomText = "Arial"
        Select Case Index
            Case 0
                .BottomTextFont.size = 6.75
                .TopTextFont.size = 6.75
                .BarcodeHeight = Text1(0)
                .QuietZone = Smallest_Zone
            Case 1
                .BottomTextFont.size = 8
                .TopTextFont.size = 8
                .BarcodeHeight = Text1(2)
                .QuietZone = Medium_Zone
            Case 2
                .BottomTextFont.size = 12
                .TopTextFont.size = 12
                .BarcodeHeight = Text1(4)
                .QuietZone = Largest_Zone
        End Select
        .Barcode = "M" & lista.ListItems(linea_lista).Text
        .BottomText = "Código: " & lista.ListItems(linea_lista).SubItems(1)
'        .Barcode = "M" & muestra_proceso
'        .BottomText = "Código: " & oMuestra.CodigoParticular(muestra_proceso)
        .TopText = "Ref:" & Left(oMuestra.getREFERENCIA_CLIENTE, 30)
    End With
    Barcode.Left = (Me.Width / 2) - (Barcode.Width / 2)
End Sub
Public Sub cargar_medidas()
    Dim oParametros As New clsParametros
    ' Pequeña
    If oParametros.Carga(parametros.ETIQUETA_PEQUEÑA, "") Then
        Text1(0) = oParametros.getVALOR
    End If
    ' Mediana
    If oParametros.Carga(parametros.ETIQUETA_MEDIANA, "") Then
        Text1(2) = oParametros.getVALOR
    End If
    ' Grande
    If oParametros.Carga(parametros.ETIQUETA_GRANDE, "") Then
        Text1(4) = oParametros.getVALOR
    End If
End Sub
Public Sub cargar_lista()
    Dim oMuestra As New clsMuestra
    Dim x As Integer
    Dim i As Integer
    For x = 1 To UBound(etiquetas, 1)
        If oMuestra.esControlEficacia(etiquetas(x)) Then
            Dim oCe_resultados As New clsCe_resultados
            Dim rs As ADODB.Recordset
            Dim encontrado As Boolean
            Set rs = oCe_resultados.Listado_por_muestra(etiquetas(x))
            If rs.RecordCount > 0 Then
                Do
                    encontrado = False
                    For i = 1 To lista.ListItems.Count
                        If lista.ListItems(i).SubItems(1) = rs("IDENTIFICACION_CANAGROSA") Then
                            encontrado = True
                        End If
                    Next
                    If Not encontrado Then
                        With lista.ListItems.Add(, , etiquetas(x))
                            .SubItems(1) = oMuestra.CodigoParticular(etiquetas(x)) & " (" & rs("IDENTIFICACION_CANAGROSA") & ")"
                        End With
                    End If
                    rs.MoveNext
                Loop Until rs.EOF
            End If
        Else
            With lista.ListItems.Add(, , etiquetas(x))
                .SubItems(1) = oMuestra.CodigoParticular(etiquetas(x))
            End With
        End If
    Next
End Sub

