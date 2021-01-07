VERSION 5.00
Object = "{F4375239-2DAA-489A-9DCE-662FC9185BD6}#1.99#0"; "BarcodeWiz.dll"
Begin VB.Form frmEquipoEtiqueta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Etiqueta de Equipo"
   ClientHeight    =   3450
   ClientLeft      =   4725
   ClientTop       =   3360
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1050
   End
   Begin VB.TextBox txtResponsableTecnico 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3390
      TabIndex        =   7
      Top             =   2010
      Width           =   3270
   End
   Begin VB.TextBox txtLocalizacion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3390
      TabIndex        =   5
      Top             =   1560
      Width           =   3270
   End
   Begin VB.TextBox txtNombreEquipo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   30
      TabIndex        =   4
      Top             =   1560
      Width           =   3270
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1050
   End
   Begin VB.TextBox txtNSerie 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3390
      TabIndex        =   2
      Top             =   1110
      Width           =   3270
   End
   Begin VB.TextBox txtNumeroEquipo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   30
      TabIndex        =   0
      Top             =   1110
      Width           =   3270
   End
   Begin VB.TextBox txtCondicionesAmbientales 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   30
      TabIndex        =   6
      Top             =   2010
      Width           =   3270
   End
   Begin BARCODEWIZLibCtl.BarCodeWiz bcdEtiqueta 
      Height          =   750
      Left            =   2160
      TabIndex        =   1
      Top             =   90
      Width           =   2340
      m_scaleNumerator=   1
      m_scaleDenominator=   1
      _cx             =   4125
      _cy             =   1327
      AutoSize        =   -1  'True
      BackColor       =   16777215
      BackStyle       =   1
      Barcode         =   "1234"
      BarcodeHeight   =   1000
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
      Border          =   1
      BottomText      =   ""
      BeginProperty BottomTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      QuietZone       =   3
      ScaleMode       =   0
      StretchBarcodeText=   0   'False
      Symbology       =   0
      TopText         =   ""
      TopTextAlignment=   2
      BeginProperty TopTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WideToNarrowRatio=   0
   End
End
Attribute VB_Name = "frmEquipoEtiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarobjEquipo As clsEquipos
Private mvarstrResponsable As String
Private mvarstrLocalizacion As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdImprimir_Click()


Call ImprimirEtiquetaEquipo


End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub

Private Sub Form_Load()
bcdEtiqueta.Left = (Me.ScaleWidth / 2) - (bcdEtiqueta.Width / 2)

Call cargar_botones(Me)

Call PresentarDatos

End Sub



Public Property Get Equipo() As clsEquipos

    Set Equipo = mvarobjEquipo

End Property

Public Property Set Equipo(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Private Sub PresentarDatos()

    With mvarobjEquipo
        bcdEtiqueta.Barcode = .getID_EQUIPO
        'txtLocalizacion = .getLOCALIZACION
        txtNombreEquipo = .getNOMBRE
        txtNSerie = .getSERIE
        txtNumeroEquipo = .getID_EQUIPO
        txtLocalizacion.Text = mvarstrLocalizacion
        txtResponsableTecnico = mvarstrResponsable
        If .getCONDICIONES_AMBIENTALES <> 0 Then
            txtCondicionesAmbientales.Text = "Temp.(Cº): Min. " & CStr(.getTEMPERATURA_MIN) & " / Max. " & CStr(.getTEMPERATURA_MAX)
            txtCondicionesAmbientales.Text = "Hume.(%Hr): Min. " & CStr(.getHUMEDAD_MIN) & " / Max. " & CStr(.getHUMEDAD_MAX)
        Else
            txtCondicionesAmbientales.Text = "N/A"
        End If
    End With
    
End Sub

Public Property Get responsable() As String

    responsable = mvarstrResponsable

End Property

Public Property Let responsable(ByVal strRESPONSABLE As String)

    mvarstrResponsable = strRESPONSABLE

End Property

Public Property Get Localizacion() As String

    Localizacion = mvarstrLocalizacion

End Property

Public Property Let Localizacion(ByVal strLocalizacion As String)

    mvarstrLocalizacion = strLocalizacion

End Property

Private Sub ImprimirEtiquetaEquipo()
    Dim i As Long
    Dim strEquipos As String
    Dim booAlgunoSeleccionado As Boolean
    
On Error GoTo trataError

    ' se mira si el equipo tiene impresora de etiquetas
'    Dim oParametro As New clsParametros
'    If Not oParametro.Carga(parametros.IMPRESORA_ETIQUETAS_GRANDE, USUARIO.getUSO) Then
'        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
'        Exit Sub
'    End If
'    log ("Comienzo impresion de etiquetas de equipos")
'    Dim impresora_encontrada As Boolean
'    impresora_encontrada = False
'    For Each prnPrinter In Printers
'        If prnPrinter.DeviceName = Replace(oParametro.getVALOR, "/", "\") Then
'            Set Printer = prnPrinter
'            impresora_encontrada = True
'            Exit For
'        End If
'    Next
'    If impresora_encontrada Then

    Dim objfrm As New frmReport

        With objfrm
            Firmas.copiar_firma_responsable_tecnico
            
            .iniciar
            .informe = "Equipos\rptEquipos_Etiqueta"
            strEquipos = "{equipos.ID_EQUIPO} in [" & mvarobjEquipo.getID_EQUIPO & ".00]"
            'booAlgunoSeleccionado = False
            'strEquipos = strEquipos & CLng(lista.ListItems(i)) & ".00,"
            booAlgunoSeleccionado = True
                        
            If booAlgunoSeleccionado Then
                strEquipos = strEquipos ' Left(strEquipos, Len(strEquipos) - 1) & "]"
                .CRITERIO = strEquipos
                .imprimir = False
                .generar
                '.Visible = True
                .Show 1
            Else
                MsgBox "Debe marcar los equipos para los que desea generar etiqueta.", vbOKOnly + vbInformation, App.Title
            End If
        End With
        
        Unload objfrm
        Set objfrm = Nothing
        log ("Final impresion de etiquetas de equipos")
'    Else
'        MsgBox "No se localiza la impresora definida en el parámetro.", vbExclamation, App.Title
'    End If
    Exit Sub
    
trataError:
    MsgBox "Error al imprimir las etiquetas.", vbCritical, Err.Description
End Sub
