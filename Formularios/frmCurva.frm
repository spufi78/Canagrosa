VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmCurva 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Curva de Alveograma"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12900
   Icon            =   "frmCurva.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   12900
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flg 
      Height          =   3495
      Left            =   7230
      TabIndex        =   1
      Top             =   -30
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   3
   End
   Begin MSChart20Lib.MSChart grafico 
      Height          =   3855
      Left            =   390
      OleObjectBlob   =   "frmCurva.frx":030A
      TabIndex        =   0
      Top             =   450
      Width           =   5025
   End
   Begin VB.Image cmdCancel 
      Height          =   615
      Left            =   4860
      MouseIcon       =   "frmCurva.frx":34EA
      Picture         =   "frmCurva.frx":37F4
      Stretch         =   -1  'True
      Top             =   5070
      Width           =   1050
   End
End
Attribute VB_Name = "frmCurva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Dim fichero As String
    Dim fechas As String
    Dim HORA As String
    Dim EMPLEADO As String
    Dim MUESTRA As String
    Dim CODIGO As String
    Dim linea As String
    Dim val(16) As String
    Dim x As Integer
    Dim i As Integer
    Dim j As Integer
    ' Enlace1.txt
    fichero = ReadINI(App.Path + "\config.ini", "Alveograma", "ruta")
    fichero = fichero & "\enlace1.txt"
    ' Comprobar si existe el fichero
    If Dir(fichero) = "" Then
       MsgBox "No existe el fichero de curvas.", vbInformation, App.Title
       Exit Sub
    End If
    
    Open fichero For Input As #1
    Dim curvas(5, 100) As String * 3
    Dim eje(5, 100) As String * 3
    For i = 1 To 5
        For j = 1 To 100
            curvas(i, j) = ""
            eje(i, j) = ""
        Next
    Next
    For i = 1 To 5
        Line Input #1, linea
        Pos = 8
        x = 1
        Do Until Pos > Len(linea)
            eje(i, x) = Mid(linea, Pos, 3)
            curvas(i, x) = Mid(linea, Pos + 3, 3)
            Pos = Pos + 6
            x = x + 1
        Loop
    Next
    Close #1
    
    With grafico
'        .ShowLegend = True
'        .chartType = VtChChartType2dLine
'        .Title.Text = " web site end-to-end time"
'        .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "Date"
'        .Plot.Axis(VtChAxisIdY).AxisTitle.Text = "Seconds"
'        .FootnoteText = "some note"
'        .ChartData = curvas
        .ColumnCount = 1
        Dim MaxCol As Integer
        MaxCol = 0
        maxcol2 = 0
        For i = 1 To 1
          x = 1
          Do
            If x > MaxCol Then
                MaxCol = x
            End If
            If CInt(eje(i, x)) > maxcol2 Then
                maxcol2 = CInt(eje(i, x))
            End If
            x = x + 1
          Loop Until Trim(eje(i, x)) = ""
        Next
        .RowCount = maxcol2
        flg.Rows = MaxCol + 1
        x = 1
        i = 1
        Do
            flg.Row = x
            flg.Col = 1
            flg.Text = CInt(eje(i, x))
            flg.Col = 2
            flg.Text = CInt(curvas(i, x))
            x = x + 1
        Loop Until Trim(curvas(i, x)) = "" Or Trim(eje(i, x)) = ""

        For x = 1 To flg.Rows - 1
            flg.Row = x
            flg.Col = 1
            If CInt(flg.Text) = 0 Then
                .Row = 1
            Else
                .Row = CInt(flg.Text)
            End If
            flg.Col = 2
            'Ponemos los titulos de los ejes
            .RowLabel = "Valor " & x
            .Data = flg.Text
        Next x
        
'        For i = 1 To 1
'            .Column = i
'            x = 1
'            Do
'                If eje(i, x) = eje(i, x - 1) Then
'                    If x = 1 Then
'                        .Row = 1
'                    Else
'                        .Row = .Row + 1
'                    End If
'                Else
'                    If CInt(eje(i, x)) = 0 Then
'                        .Row = 1
'                    Else
'                        If CInt(eje(i, x)) > .Row Then
'                            .Row = CInt(eje(i, x))
'                        Else
'                            .Row = .Row + 1
'                        End If
'                    End If
'                End If
'                If CInt(curvas(i, x)) = 0 Then
'                    .Data = 1
'                Else
'                    .Data = CInt(curvas(i, x))
'                End If
'                .Row = x
'                .Data = x
'                x = x + 1
'            Loop Until Trim(curvas(i, x)) = "" Or Trim(eje(i, x)) = ""
'            .Data = 1
'        Next
'       .Refresh
    End With
End Sub

