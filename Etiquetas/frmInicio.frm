VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H00808080&
   Caption         =   "Cliente de Etiquetado"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2025
      Top             =   1350
   End
   Begin VB.Label lblLog 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   45
      TabIndex        =   1
      Top             =   1260
      Width           =   4380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "INICIANDO SISTEMA...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   4290
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    If App.PrevInstance = True Then
        Timer1.Enabled = False
        MsgBox "El cliente de etiquetado ya se encuentra en ejecución. Verifique la ejecución anterior.", vbInformation, App.Title
        End
    End If
    Timer1.Enabled = False
    Crear_DSN
    ' Cargar plantillas
'    If CrearConexionGlobal = False Then
'        MsgBox "Error al crear la conexión global. Contacte con mantenimiento.", vbCritical, App.Title
'        End
'    End If
    Dim rs As ADODB.Recordset
    Dim c As String
    c = "select * from etiquetas_rpt order by id_tipo"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            lblLog = "Creando plantilla : " & rs("DESCRIPCION")
            Dim mystream As New ADODB.Stream
            mystream.Type = adTypeBinary
            mystream.Open
            mystream.Write rs("RPT")
            On Error Resume Next
            Dim ruta As String
            ruta = ReadINI(App.Path + "\config.ini", "documentos", "reportes")
            MkDir ruta
            ' Crear estructura de carpetas
            ruta = ruta & "\" & rs("CARPETA")
            MkDir ruta
            Dim fichero
            fichero = ruta & "\" & rs("INFORME")
            mystream.SaveToFile fichero, adSaveCreateOverWrite
            mystream.Close
'            rs.Close
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    On Error Resume Next
    MkDir App.Path & "\firmas"
    
    Timer1.Enabled = True

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmInicio"
    
End Sub
Private Sub Timer1_Timer()
    database = ReadINI(App.Path + "\config.ini", "server", "bd")
    Unload Me
    frmPDF.Show 1
End Sub
