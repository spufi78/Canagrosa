VERSION 5.00
Object = "{F4375239-2DAA-489A-9DCE-662FC9185BD6}#1.99#0"; "BarcodeWiz.dll"
Begin VB.Form frmEquipoEtiquetaCalibracion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Etiqueta de Calibración"
   ClientHeight    =   3450
   ClientLeft      =   5145
   ClientTop       =   3225
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
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
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3540
      TabIndex        =   7
      Top             =   2010
      Width           =   3000
   End
   Begin VB.TextBox txtCondicionesAmbientales 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      TabIndex        =   6
      Top             =   2010
      Width           =   3000
   End
   Begin VB.TextBox txtLocalizacion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3540
      TabIndex        =   5
      Top             =   1560
      Width           =   3000
   End
   Begin VB.TextBox txtNombreEquipo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      TabIndex        =   4
      Top             =   1560
      Width           =   3000
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
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3540
      TabIndex        =   2
      Top             =   1110
      Width           =   3000
   End
   Begin VB.TextBox txtNumeroEquipo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      TabIndex        =   0
      Top             =   1110
      Width           =   3000
   End
   Begin BARCODEWIZLibCtl.BarCodeWiz bcdEtiqueta 
      Height          =   750
      Left            =   2160
      TabIndex        =   1
      Top             =   240
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
Attribute VB_Name = "frmEquipoEtiquetaCalibracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarobjEquipo As clsEquipos

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
'        bcdEtiqueta.Barcode = .getID_EQUIPO
'        'txtLocalizacion = .getLOCALIZACION
'        txtNombreEquipo = ""
'        txtNSerie = ""
'        txtNumeroEquipo = ""
'        txtResponsableTecnico = ""
'        txtCondicionesAmbientales.Text = ""
    End With
    
End Sub
