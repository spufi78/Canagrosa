Attribute VB_Name = "Correo"

Public Function Enviar_Mail_CDO(Para As String, asunto As String, Mensaje As String, Optional Path_Adjunto As String)
   SerVidor_SMTP = "smtp.gmail.com"
   user = "julio.gonzalez.moreno@gmail.com"
   PASSWORD = "man0lete"
   De = "BCA"
   asunto = "BCA : " & asunto
   Puerto = "465"
   Usar_Autentificacion = True
   Usar_SSL = True
  
  
    ' Variable de objeto Cdo.Message
    Dim Obj_Email As CDO.Message
          
    
    ' Crea un Nuevo objeto CDO.Message
    Set Obj_Email = New CDO.Message
    
    ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
      del servidor o su dirección IP )
    Obj_Email.Configuration.Fields(cdoSMTPServer) = SerVidor_SMTP
    
    Obj_Email.Configuration.Fields(cdoSendUsingMethod) = 2
    
    ' Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usan los puertos _
      465 o  el puerto 587 ( este último me dio error )
    
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(Puerto)

    
    ' Indica el tipo de autentificación con el servidor de correo _
     El valor 0 no requiere autentificarse, el valor 1 es con autentificación
    Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/" & _
                "configuration/smtpauthenticate") = Abs(Usar_Autentificacion)
    
    
    
        ' Tiempo máximo de espera en segundos para la conexión
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

    
    ' Configura las opciones para el login en el SMTP
    If Usar_Autentificacion Then

    ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la dirección de correro _
     mas el @gmail.com )
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendusername") = user

    ' Password de la cuenta
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PASSWORD

    ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Usar_SSL
    
    End If
    

    ' *********************************************************************************
    ' Estructura del mail
    '**********************************************************************************
    
    ' Dirección del Destinatario
    Obj_Email.To = Para
    
    ' Dirección del remitente
    Obj_Email.From = De
    
    ' Asunto del mensaje
    Obj_Email.Subject = asunto
    
    ' Cuerpo del mensaje
    Obj_Email.TextBody = Mensaje
    
    'Ruta del archivo adjunto
    
    If Path_Adjunto <> vbNullString Then
        Obj_Email.AddAttachment (Path_Adjunto)
    End If
    
    ' Actualiza los datos antes de enviar
    Obj_Email.Configuration.Fields.Update
    
    On Error Resume Next
    ' Envía el email
    Obj_Email.Send
    
    
    If Err.Number = 0 Then
       Enviar_Mail_CDO = True
    Else
       MsgBox Err.Description, vbCritical, " Error al enviar el correo "
    End If
    
    ' Descarga la referencia
    If Not Obj_Email Is Nothing Then
        Set Obj_Email = Nothing
    End If
    
    On Error GoTo 0

End Function


Public Function genera_correo_outlook(mailto As String, asunto As String, cuerpo As String, destino_documento As String, manejador As Long)
    If Trim(mailto) <> "" Then
        enviar_correo mailto, "", "", True, cuerpo, asunto, destino_documento
    Else
        enviar_correo "Introduzca destinatario", "", "", True, asunto, cuerpo, destino_documento
    End If
End Function

Public Sub enviar_correo(sender As String, ccto As String, bccto As String, DisplayMsg As Boolean, mBody As String, mSubject As String, Optional AttachmentPath)
       Dim objOutlook As Outlook.Application
       Dim objOutlookMsg As Outlook.MailItem
       Dim objOutlookRecip As Outlook.Recipient
       Dim objOutlookAttach As Outlook.Attachment
       ' Create the Outlook session.
   On Error GoTo enviar_correo_Error

       Set objOutlook = CreateObject("Outlook.Application")
       ' Create the message.
       Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
       
       With objOutlookMsg
           ' Add the To recipient(s) to the message.
           Set objOutlookRecip = .Recipients.Add(sender)
           objOutlookRecip.Type = olTo

           ' Add the CC recipient(s) to the message.
           If ccto <> "" Then
            Set objOutlookRecip = .Recipients.Add(ccto)
            objOutlookRecip.Type = olCC
           End If

          ' Add the BCC recipient(s) to the message.
          If bccto <> "" Then
           Set objOutlookRecip = .Recipients.Add(bccto)
           objOutlookRecip.Type = olBCC
          End If

          ' Set the Subject, Body, and Importance of the message.
          .Subject = mSubject
          .body = mBody
          .Importance = olImportanceHigh  'High importance

          ' Add attachments to the message.
          If Not IsMissing(AttachmentPath) Then
              Dim ficheros() As String
              ficheros = Split(AttachmentPath, ";")
              Dim i As Integer
              For i = LBound(ficheros) To UBound(ficheros)
'                  Set objOutlookAttach = .Attachments.Add(AttachmentPath)
                If ficheros(i) <> "" Then
                    If Dir(ficheros(i)) <> "" Then
                      Set objOutlookAttach = .Attachments.Add(ficheros(i))
                    End If
                End If
              Next
          End If

          ' Resolve each Recipient's name.
          For Each objOutlookRecip In .Recipients
              objOutlookRecip.Resolve
          Next

          ' Should we display the message before sending?
          If DisplayMsg Then
              .Display
          Else
              .Save
              .Send
          End If
       End With
'       Set objOutlook = Nothing

   On Error GoTo 0
   Exit Sub

enviar_correo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure enviar_correo of Módulo Correo"
End Sub



