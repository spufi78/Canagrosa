VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#13.2#0"; "Codejock.Calendar.v13.2.1.ocx"
Begin VB.Form frmCalendario 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Calendario de Finalización de Ensayos de Tiempo"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCalendario.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9975
   ScaleWidth      =   11640
   WindowState     =   1  'Minimized
   Begin XtremeCalendarControl.CalendarControl CalendarControl 
      Height          =   8745
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   11445
      _Version        =   851970
      _ExtentX        =   20188
      _ExtentY        =   15425
      _StockProps     =   64
      ViewType        =   2
      ShowCaptionBar  =   -1  'True
   End
   Begin VB.Frame frmOpciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   60
      TabIndex        =   1
      Top             =   8880
      Width           =   6825
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   870
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1050
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   870
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker finicio 
         Height          =   330
         Left            =   1650
         TabIndex        =   3
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   60227585
         CurrentDate     =   38002
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha Inicio Tareas"
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdMinimizar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Minimizar"
      Height          =   870
      Left            =   10530
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   1050
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    nuevo_evento Date
End Sub
Private Sub nuevo_evento(fecha As String)
    frmMensaje_Detalle.PK = 0
    frmMensaje_Detalle.fdesde = fecha
    frmMensaje_Detalle.fhasta = fecha
    frmMensaje_Detalle.txttexto(3) = 0
    frmMensaje_Detalle.Show 1
End Sub
Private Sub CalendarControl_DblClick()
    Dim HitTest As CalendarHitTestInfo
    Set HitTest = CalendarControl.ActiveView.HitTest

    Dim Events As CalendarEvents
    If Not HitTest.HitCode = xtpCalendarHitTestUnknown Then
     '   Set Events = CalendarControl.DataProvider.RetrieveDayEvents(HitTest.ViewDay.Date)
    End If

    If HitTest.ViewEvent Is Nothing Then
        nuevo_evento HitTest.ViewDay.Date
    Else
'        If HitTest.ViewEvent.Event.Categories(0) = 6 Then
'        MsgBox HitTest.ViewEvent.Event
'        MsgBox HitTest.ViewEvent.Event.Label
'            gmuestra = HitTest.ViewEvent.Event.Label
'            frmVerMuestra.Show 1
'            gmuestra = 0
'        Else
            frmMensaje_Detalle.PK = HitTest.ViewEvent.Event.Label
            frmMensaje_Detalle.Show 1
'        End If
'        ModifyEvent HitTest.ViewEvent.Event
    End If
End Sub

Private Sub cmdImprimir_Click()
    CalendarControl.PrintPreviewOptions.Title = "Canagrosa"
    
    CalendarControl.PrintOptions.Header.Font.bold = True
    CalendarControl.PrintOptions.Header.TextLeft = "Canagrosa"
    CalendarControl.PrintOptions.Header.TextCenter = "Listado de tareas"
'    CalendarControl.PrintOptions.Header.TextRight = "Página 1 de 1 "
    
    CalendarControl.PrintOptions.Footer.TextLeft = "Fecha: " & DateValue(Now) & " Hora: " & TimeValue(Now)
'    CalendarControl.PrintOptions.Footer.TextCenter = "Canagrosa, " & vbLf & " Listado de Tareas "
    CalendarControl.PrintOptions.Footer.TextRight = "Página 1 de 1 "
    
    CalendarControl.PrintPreviewExt True, 200, 200, 800, 600
End Sub

Private Sub cmdMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub finicio_Change()
    cargar_eventos
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    finicio = Date - 30
    cargar_eventos
    CalendarControl.Options.EnableInPlaceEditEventSubject_ByMouseClick = False

'    CalendarControl.ReadOnlyMode = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Cancel = True
'    Me.WindowState = vbMinimized
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim nHeight As Long, LabelWidth As Long
    CalendarControl.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 1100

'    cmdAnadir.Left = Me.ScaleWidth - cmdAnadir.Width - cmdMinimizar.Width - 200
'    cmdAnadir.Top = Me.ScaleHeight - cmdAnadir.Height - 100
    
    cmdminimizar.Left = Me.ScaleWidth - cmdminimizar.Width - 100
    cmdminimizar.top = Me.ScaleHeight - cmdminimizar.Height - 100

    frmOpciones.Left = 100
    frmOpciones.top = Me.ScaleHeight - frmOpciones.Height - 50

End Sub

Public Sub cargar_eventos()
    
    CalendarControl.DataProvider.RemoveAllEvents
    
    Dim rs As ADODB.Recordset
    Dim consulta As String
'    consulta = "select b.ID_MUESTRA, tm.CODIGO, b.ID_PARTICULAR,a.DURACION_FECHA_HASTA, a.DURACION_HORA_HASTA, b.referencia_cliente " & _
'               "  from ce_recepcion a, muestras b, tipos_muestra tm " & _
'               " where a.MUESTRA_ID = b.ID_MUESTRA " & _
'               "   and b.CERRADA = 0 and b.ANULADA = 0 " & _
'               "   and duracion_fecha_hasta <> '' " & _
'               "   and duracion_fecha_hasta <> '01-01-1900' " & _
'               "   and b.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA " & _
'               " order by b.ID_MUESTRA"
'    Set rs = datos_bd(consulta)
'    If rs.RecordCount > 0 Then
'        Do
'            addevent rs(0), rs(1) & "-" & rs(2) & " (" & rs(5) & ")", "", rs(3), rs(4), 0, rs(3), 6
'            rs.MoveNext
'        Loop Until rs.EOF
'    End If

    Dim oMensaje As New clsMensajes
    Set rs = oMensaje.Listado_completo(finicio.Value)
    If rs.RecordCount <> 0 Then
        Do
            addevent rs(0), rs(1), rs(7), rs(5), rs(8), rs(11), rs(6), rs(10)
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oMensaje = Nothing
    
    Set rs = Nothing
End Sub

Private Sub addevent(ID As Long, texto As String, body As String, fecha_inicio As String, HORA_INICIO As String, duracion_minutos As Long, fecha_fin As String, categoria As Integer)
    Dim NewEvent As CalendarEvent, Recurrence As CalendarRecurrencePattern
    Set NewEvent = CalendarControl.DataProvider.CreateEvent
    
    NewEvent.Subject = texto
    NewEvent.LOCATION = ""
    NewEvent.body = body
    NewEvent.ReminderSoundFile = ".."
    NewEvent.Label = ID
    NewEvent.Categories.Add categoria
'    NewEvent.Importance = xtpCalendarImportanceHigh
'    NewEvent.MeetingFlag = True
'    NewEvent.PrivateFlag = True
    
    Set Recurrence = NewEvent.CreateRecurrence
    Recurrence.StartTime = Format(HORA_INICIO, "hh:mm:ss")
    Recurrence.DurationMinutes = duracion_minutos
    Recurrence.StartDate = CDate(fecha_inicio)
    If fecha_fin = "" Then
        Recurrence.EndDate = CDate(fecha_inicio)
    Else
        Recurrence.EndDate = CDate(fecha_fin)
    End If
'    Recurrence.Options.RecurrenceType = xtpCalendarRecurrenceWeekly
    Recurrence.Options.RecurrenceType = xtpCalendarRecurrenceDaily

    Recurrence.Options.WeeklyIntervalWeeks = 1
    Recurrence.Options.WeeklyDayOfWeekMask = xtpCalendarDayAllWeek
    NewEvent.UpdateRecurrence Recurrence
    CalendarControl.DataProvider.addevent NewEvent
End Sub
