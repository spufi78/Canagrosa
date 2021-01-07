VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaptura 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Captura de curvas de alveogramas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmCaptura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   10620
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   4365
      TabIndex        =   6
      Top             =   6885
      Width           =   825
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6420
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   9495
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6420
      Width           =   1050
   End
   Begin VB.CommandButton cmdCapturar 
      Caption         =   "Capturar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   90
      Picture         =   "frmCaptura.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6420
      Width           =   1365
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5925
      Left            =   45
      TabIndex        =   2
      Top             =   405
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   10451
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Doble Click para ver los datos de la captura"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   6360
      Width           =   4335
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Captura de curvas"
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
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   10545
   End
End
Attribute VB_Name = "frmCaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdCapturar_Click()
    On Error GoTo fallo
    If MsgBox("Va a comenzar la captura de datos. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim fichero As String
        Dim fechas As String
        Dim HORA As String
        Dim EMPLEADO As String
        Dim muestra As String
        Dim CODIGO As String
        Dim linea As String
        Dim val(16) As String
        ' Enlace.txt
        fichero = ReadINI(App.Path + "\config.ini", "Alveograma", "ruta")
        fichero = fichero & "\enlace.txt"
        ' Comprobar si existe el fichero
        If Dir(fichero) = "" Then
            MsgBox "No existe el fichero de curvas.", vbInformation, App.Title
            Exit Sub
        End If
        lista.ListItems.Clear
        Open fichero For Input As #1
        Line Input #1, linea
        Do
            val(13) = Mid(linea, 73, 5)
            val(14) = Mid(linea, 78, 5)
            val(15) = Mid(linea, 83, 5)
            ' El resultado debe ser un numero entero al dividir por 1.1
            'val(0) = Mid(linea, 7, 5)
            'val(1) = Mid(linea, 35, 5)
            val(0) = CSng(Replace(Mid(linea, 7, 5), ".", ",")) / 1.1
            val(0) = CStr(CSng(CInt(val(0)) * 1.1))
            val(1) = CSng(Replace(Mid(linea, 35, 5), ".", ",")) / 1.1
            val(1) = CStr(CSng(CInt(val(1)) * 1.1))
            val(2) = Mid(linea, 12, 3)
            val(3) = Mid(linea, 40, 3)
            val(4) = formatear(CSng(val(0)) / CSng(val(2)), 5, 2)
            val(5) = formatear(CSng(val(1)) / CSng(val(3)), 5, 2)
            val(6) = Mid(linea, 18, 5)
            val(7) = Mid(linea, 46, 5)
            val(8) = Mid(linea, 15, 3)
            val(9) = Mid(linea, 43, 3)
            val(10) = Mid(linea, 23, 6)
            val(11) = Mid(linea, 51, 6)
            val(12) = Mid(linea, 57, 3)
            muestra = Mid(linea, 5, 2)
            CODIGO = Mid(linea, 1, 4)
            Fecha = Mid(linea, 60, 8)
            HORA = Mid(linea, 68, 5)
            EMPLEADO = Mid(linea, 87, 2)
            
            With lista.ListItems.Add(, , muestra & "-" & CODIGO)
                .SubItems(1) = val(13)
                .SubItems(2) = val(14)
                .SubItems(3) = val(15)
                .SubItems(4) = val(12)
                .SubItems(5) = Left(Fecha, 2) & "/" & Mid(Fecha, 3, 2) & "/" & Right(Fecha, 4)
                .SubItems(6) = HORA
                .SubItems(7) = val(0)
                .SubItems(8) = val(2)
                .SubItems(9) = val(4)
                .SubItems(10) = val(6)
                .SubItems(11) = val(8)
                .SubItems(12) = val(10)
                .SubItems(13) = val(1)
                .SubItems(14) = val(3)
                .SubItems(15) = val(5)
                .SubItems(16) = val(7)
                .SubItems(17) = val(9)
                .SubItems(18) = val(11)
            End With
            lista.ListItems(lista.ListItems.Count).Checked = True
            linea = ""
            If EOF(1) = False Then
                Line Input #1, linea
            End If
        Loop Until linea = "" And EOF(1) = True
        Close #1
        
    End If
    Exit Sub
fallo:
    Close #1
    MsgBox "Error al capturar los datos del fichero.", vbCritical, Err.Description
End Sub

Private Sub cmdOk_Click()
    On Error GoTo fallo
    If MsgBox("Va a introducir los datos de la Captura. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        ' Captura
        Dim alveo As Long
        Dim oalveoval As New clsAlveograma_valores
        Dim oalveo As New clsAlveogramas
        Dim omuestra As New clsMuestra
        Dim oDeter As New clsDeterminaciones
        Dim odd As New clsDatos_determinaciones
        Dim rs As New ADODB.Recordset
        Dim i As Integer
        Dim CODIGO As String
        Dim pos As Integer
        Dim particular As Integer
        Dim muestra As Long
        Dim DETERMINACION As Long
        For i = 1 To lista.ListItems.Count
         If lista.ListItems(i).Checked = True Then
            CODIGO = lista.ListItems(i)
            pos = InStr(1, CODIGO, "-", vbTextCompare)
            particular = Mid(CODIGO, pos + 1, Len(CODIGO))
            consulta = "select * from tipos_muestra where codigo = '" & Trim(Mid(CODIGO, 1, pos - 1)) & "'"
            Set rs = datos_bd(consulta)
            consulta = "select * from muestras where tipo_muestra_id = " & rs("id_tipo_muestra") & " and id_particular = " & particular & " and anno = " & Year(Date)
            Set rs = datos_bd(consulta)
            If rs.RecordCount = 0 Then
                MsgBox "No existe la muestra " & lista.ListItems(i), vbExclamation, App.Title
            Else
                muestra = rs("id_muestra")
                consulta = "select * from determinaciones where muestra_id = " & muestra & " and tipo_determinacion_id = " & CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma"))
                Set rs = datos_bd(consulta)
                If rs.RecordCount = 0 Then
                    MsgBox "No existe la determinacion de alveograma para la muestra " & lista.ListItems(i), vbExclamation, App.Title
                Else
                  DETERMINACION = rs("id_determinacion")
                  consulta = "select * from alveogramas where muestra_id = " & muestra & " and determinacion_id = " & DETERMINACION
                  Set rs = datos_bd(consulta)
                  If rs.RecordCount <> 0 Then
                    MsgBox "Ya existen valores para la muestra " & lista.ListItems(i), vbExclamation, App.Title
                  Else
                    With oalveo
                     .setDETERMINACION_ID = DETERMINACION
                     .setMUESTRA_ID = muestra
                     .setTEMPERATURA = Replace(lista.ListItems(i).SubItems(1), ",", ".")
                     .setHUMEDAD_AIRE = Replace(lista.ListItems(i).SubItems(2), ",", ".")
                     .setHUMEDAD_HARINA = Replace(lista.ListItems(i).SubItems(3), ",", ".")
                     .setINDICE_DEGRADACION = Trim(Replace(lista.ListItems(i).SubItems(4), ",", "."))
                     .setFECHA = Format(lista.ListItems(i).SubItems(5), "yyyy-mm-dd")
                     alveo = .InsertarAlveograma
                    End With
                    ' Captura_Valores
                    With oalveoval
                    ' NORMAL
                     .setALVEOGRAMA_ID = alveo
                     .setTENACIDAD = Replace(lista.ListItems(i).SubItems(7), ",", ".")
                     .setEXTENSIBILIDAD = Replace(lista.ListItems(i).SubItems(8), ",", ".")
                     .setS = Replace(lista.ListItems(i).SubItems(10), ",", ".")
                     .setW = Replace(lista.ListItems(i).SubItems(11), ",", ".")
                     .setG = Replace(lista.ListItems(i).SubItems(12), ",", ".")
                     .setDE_REPOSO = 0
                     .InsertarAlveogramaValores
                     ' REPOSO
                     .setALVEOGRAMA_ID = alveo
                     .setTENACIDAD = Replace(lista.ListItems(i).SubItems(13), ",", ".")
                     .setEXTENSIBILIDAD = Replace(lista.ListItems(i).SubItems(14), ",", ".")
                     .setS = Replace(lista.ListItems(i).SubItems(16), ",", ".")
                     .setW = Replace(lista.ListItems(i).SubItems(17), ",", ".")
                     .setG = Replace(lista.ListItems(i).SubItems(18), ",", ".")
                     .setDE_REPOSO = 1
                     .InsertarAlveogramaValores
                    End With
                    ' Datos Determinaciones
                    If odd.cargar(DETERMINACION, 331) = True Then
                        odd.setVALOR_1 = Replace(lista.ListItems(i).SubItems(4), ",", ".")
                        odd.Insertar_Valores
                    End If
                    ' Almacena determinacion (Solucion)
                    oDeter.setRESULTADO = Replace(lista.ListItems(i).SubItems(17), ",", ".")
                    oDeter.setFECHA = Format(lista.ListItems(i).SubItems(5), "yyyy-mm-dd")
                    oDeter.setHORA = Format(lista.ListItems(i).SubItems(6), "hh:mm")
                    oDeter.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                    oDeter.InsertarSolucion (DETERMINACION)
                End If
            End If
          End If
         End If
        Next
        MsgBox "Captura de datos de alveogramas terminada.", vbInformation, App.Title
        Unload Me
    End If
    Exit Sub
fallo:
    MsgBox "Error al insertar los datos de la capturas(granulometria)", vbCritical, Err.Description
End Sub

Private Sub Command1_Click()
    frmAlveograma_Captura.Show 1
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
    cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    Call cabecera
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Análisis", 1000, lvwColumnLeft)
        .Tag = "Análisis"
    End With
    With lista.ColumnHeaders.Add(, , "Temperatura", 1000, lvwColumnCenter)
        .Tag = "Temperatura"
    End With
    With lista.ColumnHeaders.Add(, , "H.Aire", 1000, lvwColumnCenter)
        .Tag = "H.Aire"
    End With
    With lista.ColumnHeaders.Add(, , "H.Harina", 1000, lvwColumnCenter)
        .Tag = "H.Harina"
    End With
    With lista.ColumnHeaders.Add(, , "Indice", 1000, lvwColumnCenter)
        .Tag = "Indice"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Hora", 1000, lvwColumnCenter)
        .Tag = "Hora"
    End With
    With lista.ColumnHeaders.Add(, , "P (N)", 1000, lvwColumnCenter)
        .Tag = "P (N)"
    End With
    With lista.ColumnHeaders.Add(, , "L (N)", 1000, lvwColumnCenter)
        .Tag = "L (N)"
    End With
    With lista.ColumnHeaders.Add(, , "P/L (N)", 1000, lvwColumnCenter)
        .Tag = "P/L (N)"
    End With
    With lista.ColumnHeaders.Add(, , "S (N)", 1000, lvwColumnCenter)
        .Tag = "S (N)"
    End With
    With lista.ColumnHeaders.Add(, , "W (N)", 1000, lvwColumnCenter)
        .Tag = "W (N)"
    End With
    With lista.ColumnHeaders.Add(, , "G (N)", 1000, lvwColumnCenter)
        .Tag = "G (N)"
    End With
    With lista.ColumnHeaders.Add(, , "P (R)", 1000, lvwColumnCenter)
        .Tag = "P (R)"
    End With
    With lista.ColumnHeaders.Add(, , "L (R)", 1000, lvwColumnCenter)
        .Tag = "L (R)"
    End With
    With lista.ColumnHeaders.Add(, , "P/L (R)", 1000, lvwColumnCenter)
        .Tag = "P/L (R)"
    End With
    With lista.ColumnHeaders.Add(, , "S (R)", 1000, lvwColumnCenter)
        .Tag = "S (R)"
    End With
    With lista.ColumnHeaders.Add(, , "W (R)", 1000, lvwColumnCenter)
        .Tag = "W (R)"
    End With
    With lista.ColumnHeaders.Add(, , "G (R)", 1000, lvwColumnCenter)
        .Tag = "G (R)"
    End With

End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        frmVerAlveograma.Show 1
    End If
End Sub
