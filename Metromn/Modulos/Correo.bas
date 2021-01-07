Attribute VB_Name = "Correo"
Public Function Enviar_Mail_CDO(Para As String, ASUNTO As String, Mensaje As String, Optional Path_Adjunto As String)
   
   SerVidor_SMTP = "smtp.1and1.es"
   User = "geslab@canagrosa.com"
   PassWord = "canagrosa20"
   De = "Geslab: Canagrosa LAB&SERVICES S.L. <geslab@canagrosa.com>"
   ASUNTO = "CANAGROSA : " & ASUNTO
   Puerto = "25"
   Usar_Autentificacion = True
   Usar_SSL = False
   
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
        ("http://schemas.microsoft.com/cdo/configuration/sendusername") = User

    ' Password de la cuenta
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PassWord

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
'    Obj_Email.sender = nombre
    ' Asunto del mensaje
    Obj_Email.Subject = ASUNTO
    
    ' Cuerpo del mensaje
    Obj_Email.TextBody = Mensaje
'    Obj_Email.HTMLBody = mensaje
    
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
       Enviar_Mail_CDO = False
       MsgBox Err.Description, vbCritical, " Error al enviar el correo "
    End If
    
    ' Descarga la referencia
    If Not Obj_Email Is Nothing Then
        Set Obj_Email = Nothing
    End If
    
    On Error GoTo 0

End Function




