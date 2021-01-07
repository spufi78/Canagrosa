Attribute VB_Name = "SendMail"
' Send email using Outlook Express (IE4 or Higher)
' Developed by: Guillermo F. Zambianchi - gzambianchi@hotmail.com
' Based upon code and ideas from Bruno Paris
' Bruno's Code only sends a TXT as Attachs
' This code allows to attach any kind of file (zip/doc/xls/etc)

' Enviar email usando Outlook Express
' Sobre código original de Bruno Paris
' El original sólo enviaba TXTs como adjuntos
' este código envía cualquier tipo de archivo (zip/doc/xls/etc)

Option Explicit
' API declarations
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
'
Private Const bnd = "_NextPart_001_"     ' boundary
Private Const charset = "windows-1252"   'or "iso-8859-2"   or "windows-1250"


Public Sub SendEmail(Dest As String, ccDest As String, _
        bccDest As String, Asunto As String, Mensaje As String, _
        iAdjuntos As Integer, arAdjunto() As String, Hwnd As Long, Optional aBandejaSalida As Boolean, Optional fichero As Boolean)
' ---------------------------------------------------------------------------------
' Parametros / Parameters:
    ' Dest: Destinatarios - Recipients To:
    ' ccDest: Dest Con Copia - Cco:
    ' bccDest: Destinatarios con Copia Oculta - Bco:
    ' Asunto: Tema - Subject
    ' Mensaje: Texto del mensaje - Message Text
    ' iAdjuntos: Cantidad de Ajuntos - Number of Attachments
    ' arAdjunto: Matriz con los Adjuntos - Array with Attachment's names
    ' hwnd : handle del formulario que la invoca - Caller form's handle
    ' aBandejaSalida: Enviar directmente a Bandeja de Salida - Send to Outbox
' ---------------------------------------------------------------------------------

    Dim i As Integer, x As Integer
    Dim iRet As Long, NumFile%, FileName$
    Dim att As String
    On Error GoTo SendErr
    NumFile% = FreeFile
    ' Abrir Arch Temporario - Open temporary file
    FileName$ = getapppath() & "~temp.eml"
    Open FileName$ For Output As #NumFile%
    ' Cambiar cursor - Change Cursor shape
    Screen.MousePointer = vbHourglass
    DoEvents
    ' escribir el arch. temp. - Write to temporary file
'------------------------------------------------
    ' Destinatarios - Recipients
    If Len(Dest) > 1 Then Print #NumFile%, "To: " & Dest
    If Len(ccDest) > 1 Then Print #NumFile%, "Cc: " & ccDest
    If Len(bccDest) > 1 Then Print #NumFile%, "Bcc: " & bccDest
'------------------------------------------------
    ' Asunto - Subject
    Print #NumFile%, "Subject: " & Asunto
'------------------------------------------------
    Print #NumFile%, "MIME-Version: 1.0" & vbCrLf & "Content-Type: multipart/mixed;"
    Print #NumFile%, vbTab & "boundary=""" & bnd & """"
    Print #NumFile%, "X-Unsent: 1"
    Print #NumFile%,
    Print #NumFile%,
    Print #NumFile%, "--" & bnd
    Print #NumFile%, "Content-Type: text/plain;"
    Print #NumFile%, vbTab & "charset=""" & charset & """"
    Print #NumFile%, "Content-Transfer-Encoding: 7bit"
    Print #NumFile%,
'------------------------------------------------
    ' Texto mensaje - Msg body
    Print #NumFile%, Mensaje
    Print #NumFile%,
'------------------------------------------------
    ' Agregar Adjuntos - Add Attachs
    If fichero = True Then
        For i = 0 To iAdjuntos - 1
            Dim nArch%, linea As String * 76, Vez As Long, LineLength As Integer
            Dim Largo As Long, Veces As Long, NombreAdjunto As String
            LineLength = 76
            
            att = arAdjunto(i)
            NombreAdjunto = Right(att, Len(att) - InStrRev(att, "\"))
            Print #NumFile%, "--" & bnd
            Print #NumFile%, "Content-Type: text/plain;"  ' Funciona con ->zip/doc/xls,etc <-works with
            Print #NumFile%, vbTab & "name=""" & att & """"
            Print #NumFile%, "Content-Transfer-Encoding: 7bit"
            Print #NumFile%, "Content-Disposition: attachment;"
            Print #NumFile%, vbTab & "filename=""" & NombreAdjunto & """"
            Print #NumFile%,
            Close #NumFile%
            Open FileName$ For Binary Access Write As #NumFile%
            nArch% = FreeFile
            Seek #NumFile%, LOF(NumFile%)
            Open att For Binary Access Read As #nArch%
            Largo = LOF(nArch%)
            Veces = Int(Largo / LineLength)
            For x = 1 To Veces
                linea = Input(LineLength, nArch%)
                Put #NumFile%, , linea
            Next
            linea = Input(Largo - Veces * LineLength, nArch%)
            Put #NumFile%, , linea
            Close #nArch%
            Close #NumFile%
            Open FileName$ For Append As #NumFile%
            Seek #NumFile%, LOF(NumFile%)
            Print #NumFile%,
        Next i
    End If
'------------------------------------------------
   Print #NumFile%, "--" & bnd & "--"
    ' Cerrar arch Temp - Close temporary file
   Close #NumFile%
   Reset
     
   ' Estperar / wait 1-2 seconds
   For i = 1 To 10
       Sleep (100)
       DoEvents
   Next

'------------------------------------------------
    ' Abrir con cliente de correo predeterminado
    ' open with default e-mail program
    iRet = ShellExecute(Hwnd, "Open", _
      FileName$, _
      "", "c:\", SW_SHOWNORMAL)
    ' Si el Cliente de correo predeterminado es Outlook Express
    ' muestra la ventana de "Nuevo Mensaje"
    ' if default e-mail program is Outlook Express,
    ' this will show "New Message" window
'------------------------------------------------

    DoEvents
    Sleep (200)
    On Error Resume Next
    Dim numrtr As Integer
rtr:
    numrtr = numrtr + 1
    DoEvents
    Sleep (100)
    ' Establecer foco en la nueva ventana de Outlook Express
    ' set focus to Outlook Express window
    AppActivate Asunto
    If Err <> 0 Then
        Err = 0
        ' Si falla esperar y reintentar
        ' if not succesfull, wait and retry
        If numrtr < 100 Then GoTo rtr
    Else
        ' Enviar a la Bandeja de Salida
        '  send to outbox
        If aBandejaSalida Then SendKeys "^~"
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
SendErr:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
    Resume Next ' Exit Sub
End Sub

Function getapppath() As String
' returns path
   Dim sTemp As String
   sTemp = App.Path
   If Right$(sTemp, 1) <> "\" Then sTemp = sTemp & "\"
   getapppath = sTemp
End Function

