VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInternetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Clase para obtener la fecha y hora exactas mediante internet"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'clase para obtener la fecha de un time-server de internet y/o actualizar la fecha del sistema
'programada por José Rubí
'
Private WithEvents wsck As Winsock
Attribute wsck.VB_VarHelpID = -1
Private FechaUTC As Date
Private tzZona As TIME_ZONE_INFORMATION

'variables locales para almacenar los valores de las propiedades
Private mvarFecha As String
Private mvarHora As String
Private mvarFechayHora As Date
Private mvarCodigoError As Long
Private mvarTextoError As String
Private mvarHorarioVerano As Boolean
'
Private Type SYSTEMTIME
         wYear As Integer
         wMonth As Integer
         wDayOfWeek As Integer
         wDay As Integer
         wHour As Integer
         wMinute As Integer
         wSecond As Integer
         wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
         Bias As Long
         StandardName(32) As Integer
         StandardDate As SYSTEMTIME
         StandardBias As Long
         DaylightName(32) As Integer
         DaylightDate As SYSTEMTIME
         DaylightBias As Long
End Type
Const TIME_ZONE_ID_DAYLIGHT = 2

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function SetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME) As Long



Public Property Get HorarioVerano() As Boolean
Attribute HorarioVerano.VB_Description = "Nos informa de si nuestro ordenador está en horario de verano o no lo está."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HorarioVerano
    HorarioVerano = mvarHorarioVerano
End Property




Public Property Let FechayHora(ByVal vData As Date)
On Error GoTo FechayHora_Err

mvarFechayHora = vData
FechaUTC = DateAdd("n", tzZona.Bias, mvarFechayHora)
'si estoy en horario de verano debo quitar una hora
If mvarHorarioVerano Then FechaUTC = DateAdd("h", -1, FechaUTC)
Exit Property
FechayHora_Err:
    mvarCodigoError = Err.Number
    mvarTextoError = Err.Description
    mvarFechayHora = Empty
    FechaUTC = Empty
End Property







Public Sub ActualizarFechaSistema()
Attribute ActualizarFechaSistema.VB_Description = "Modifica la fecha y hora del sistema con el valor de la propiedad FechayHora."
Dim systime As SYSTEMTIME
Dim retval As Long

If IsEmpty(FechaUTC) Then FechaUTC = DateAdd("n", tzZona.Bias, Now())
systime.wYear = Year(FechaUTC)
systime.wMonth = Month(FechaUTC)
systime.wDay = Day(FechaUTC)
systime.wHour = Hour(FechaUTC)
systime.wMinute = Minute(FechaUTC)
systime.wSecond = Second(FechaUTC)
retval = SetSystemTime(systime)
End Sub



Public Property Get TextoError() As String
Attribute TextoError.VB_Description = "Devuelve la descripción del error si no se pudo obtener la fecha y la hora."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.TextoError
    TextoError = mvarTextoError
End Property





Public Property Get CodigoError() As Long
Attribute CodigoError.VB_Description = "Devuelve el código de error si no se pudo obtener la fecha y la hora."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CodigoError
    CodigoError = mvarCodigoError
End Property



Public Sub ObtenerFechayHora(ControlSocket As Object)
Attribute ObtenerFechayHora.VB_Description = "Asigna el control winsock que vamos a utilizar y se conecta a internet para informar de la fecha y la hora actuales."
On Error GoTo Obtener_Err

mvarFecha = ""
mvarHora = ""
mvarFechayHora = Empty
mvarCodigoError = -1
Set wsck = ControlSocket
If wsck.State <> sckClosed Then wsck.Close
wsck.Protocol = sckTCPProtocol
wsck.Connect "129.6.15.28", 13
'wsck.Connect "129.6.15.128", 13
Exit Sub
Obtener_Err:
    mvarCodigoError = Err.Number
    mvarTextoError = Err.Description
End Sub



Public Property Get FechayHora() As Date
Attribute FechayHora.VB_Description = "Devuelve la fecha y la hora actuales, en un date."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FechayHora
    FechayHora = mvarFechayHora
End Property




Public Property Get HORA() As String
Attribute HORA.VB_Description = "Devuelve la hora actual, en un string."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Hora
    HORA = mvarHora
End Property





Public Property Get fecha() As String
Attribute fecha.VB_Description = "Devuelve la fecha actual, en un string."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Fecha
    fecha = mvarFecha
End Property



Private Sub Class_Initialize()
Dim res As Long
'obtengo la información de la zona horaria
res = GetTimeZoneInformation(tzZona)
mvarHorarioVerano = (res = TIME_ZONE_ID_DAYLIGHT)
End Sub

Private Sub Class_Terminate()
Set wsck = Nothing
End Sub


Private Sub wsck_Close()
wsck.Close
End Sub

Private Sub wsck_DataArrival(ByVal bytesTotal As Long)
Dim aux As String
On Error Resume Next
wsck.GetData aux, vbString
mvarTextoError = aux
mvarCodigoError = 0
mvarFecha = Mid(aux, 14, 2) & "/" & Mid(aux, 11, 2) & "/20" & Mid(aux, 8, 2)
mvarHora = Mid(aux, 17, 8)
FechaUTC = CDate(mvarFecha & " " & mvarHora)
mvarFechayHora = DateAdd("n", -tzZona.Bias, FechaUTC)
'Debemos tener encuenta si estamos en horario de verano y el servidor no lo está
'y viceversa
If mvarHorarioVerano Then
    If Mid(aux, 26, 2) <> "50" Then
        mvarFechayHora = DateAdd("h", 1, mvarFechayHora)
    End If
Else
    If Mid(aux, 26, 2) = "50" Then
        mvarFechayHora = DateAdd("h", -1, mvarFechayHora)
    End If
End If
mvarHora = Format(mvarFechayHora, "hh:nn:ss")
End Sub


Private Sub wsck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
mvarCodigoError = Number
mvarTextoError = Description
CancelDisplay = True
wsck.Close
End Sub


