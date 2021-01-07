VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPlasma_CopiaResultados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   Icon            =   "frmPlasma_CopiaResultados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11790
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7695
      Width           =   1050
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todos"
      Height          =   330
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7695
      Width           =   1230
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todos"
      Height          =   330
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7695
      Width           =   1455
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10665
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7695
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7305
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   12885
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2025
      Top             =   8010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlasma_CopiaResultados.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlasma_CopiaResultados.frx":712C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlasma_CopiaResultados.frx":7A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlasma_CopiaResultados.frx":82E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlasma_CopiaResultados.frx":8BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlasma_CopiaResultados.frx":9494
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlasma_CopiaResultados.frx":FCF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlasma_CopiaResultados.frx":1018D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlasma_CopiaResultados.frx":10623
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ATENCIÓN, SE COPIARAN, EQUIPOS, REACTIVOS, PREPARACIÓN METALOGRÁFICA, BOND COAT, TOP COAT Y TEMPERATURA."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   2925
      TabIndex        =   6
      Top             =   7785
      Width           =   6225
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marque las muestras a las que desea copiar los resultados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11745
   End
End
Attribute VB_Name = "frmPlasma_CopiaResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdok_Click()
    Dim i As Integer
    ' Contar marcardos
    Dim marcado As Boolean
   On Error GoTo cmdok_Click_Error

    marcado = False
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            marcado = True
        End If
    Next
    If Not marcado Then
        MsgBox "Marque alguna muestra a la que copiar los resultados.", vbExclamation, App.Title
        Exit Sub
    End If
    ' Mensaje
    Dim ID_MUESTRA As Long
    If MsgBox("Va a copiar los resultados a las muestrar marcadas.¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        ' Cargamos los resultados Origen
        Dim oPRE As New clsPlasma_recepcion
        Dim oPRE_Origen As New clsPlasma_recepcion
        Dim oPR As New clsPlasma_resultados
        Dim oPRBOND_Origen As New clsPlasma_resultados
        Dim oPRTOP_Origen As New clsPlasma_resultados
        Dim oPTRA As New clsPlasma_traccion
        Dim oPTRABOND As New clsPlasma_traccion
        Dim oPTRATOP As New clsPlasma_traccion
        Dim oPAlabe As New clsPlasma_alabe
        Dim oPDureza As New clsPlasma_dureza
        
        Dim existePRBOND As Boolean
        Dim existePRTOP As Boolean
        Dim existePTRABOND As Boolean
        Dim existePTRATOP As Boolean
        Dim existePALABE As Boolean
        Dim existePDUREZA As Boolean
        
        oPRE_Origen.Carga PK
        existePRBOND = oPRBOND_Origen.Carga(PK, 1)
        existePRTOP = oPRTOP_Origen.Carga(PK, 2)
        existePTRABOND = oPTRABOND.Carga(PK, 1)
        existePTRATOP = oPTRATOP.Carga(PK, 2)
        existePALABE = oPAlabe.Carga(PK)
        existePDUREZA = oPDureza.Carga(PK)
        
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                ID_MUESTRA = lista.ListItems(i).SubItems(6)
                ' PLASMA_RECEPCION
                With oPRE
                    .setMP = oPRE_Origen.getMP
                    .setMP_FECHA = Format(oPRE_Origen.getMP_FECHA, "yyyy-mm-dd")
                    .setMP_USUARIO_ID = oPRE_Origen.getMP_USUARIO_ID
                    .setMP_PASS = oPRE_Origen.getMP_PASS
                    .setMACRO_DUREZA_T1 = oPRE_Origen.getMACRO_DUREZA_T1
                    .setMACRO_DUREZA_T2 = oPRE_Origen.getMACRO_DUREZA_T2
                    .setMICRO_DUREZA_T1 = oPRE_Origen.getMICRO_DUREZA_T1
                    .setMICRO_DUREZA_T2 = oPRE_Origen.getMICRO_DUREZA_T2
                    .setRESULT = oPRE_Origen.getRESULT
                    .ModificarCopia ID_MUESTRA
                End With
                ' Cargamos las fichas del plasma de destino para ver que ensayos hay que realizar
                oPRE.Carga ID_MUESTRA
                Dim oPPRO As New clsPlasma_procesos
                oPPRO.Carga oPRE.getPROCESO_ID
                Dim oFichaBond As New clsPlasma_ficha
                Dim oFichaTop As New clsPlasma_ficha
                oFichaBond.Carga oPPRO.getBOND_COAT_FICHA_ID
                oFichaTop.Carga oPPRO.getTOP_COAT_FICHA_ID
                
                ' PLASMA_RESULTADOS
                With oPR
                    If existePRBOND Then
                        ' BOND
                        .setMUESTRA_ID = ID_MUESTRA
                        .setTIPO = 1
                        .setBATCH = oPRBOND_Origen.getBATCH
                        If oFichaBond.getMICROESTRUCTURA <> 0 Then
                            .setMICROESTRUCTURA1 = oPRBOND_Origen.getMICROESTRUCTURA1
                            .setMICROESTRUCTURA2 = oPRBOND_Origen.getMICROESTRUCTURA2
                            .setMICROESTRUCTURA3 = oPRBOND_Origen.getMICROESTRUCTURA3
                            .setMICROESTRUCTURA4 = oPRBOND_Origen.getMICROESTRUCTURA4
                            .setMICROESTRUCTURA5 = oPRBOND_Origen.getMICROESTRUCTURA5
                            .setMICROESTRUCTURA6 = oPRBOND_Origen.getMICROESTRUCTURA6
                            .setMICROESTRUCTURA1_R = oPRBOND_Origen.getMICROESTRUCTURA1_R
                            .setMICROESTRUCTURA2_R = oPRBOND_Origen.getMICROESTRUCTURA2_R
                            .setMICROESTRUCTURA3_R = oPRBOND_Origen.getMICROESTRUCTURA3_R
                            .setMICROESTRUCTURA4_R = oPRBOND_Origen.getMICROESTRUCTURA4_R
                            .setMICROESTRUCTURA5_R = oPRBOND_Origen.getMICROESTRUCTURA5_R
                            .setMICROESTRUCTURA6_R = oPRBOND_Origen.getMICROESTRUCTURA6_R
                            .setMICROESTRUCTURA1_VALOR = oPRBOND_Origen.getMICROESTRUCTURA1_VALOR
                            .setMICROESTRUCTURA2_VALOR = oPRBOND_Origen.getMICROESTRUCTURA2_VALOR
                            .setMICROESTRUCTURA3_VALOR = oPRBOND_Origen.getMICROESTRUCTURA3_VALOR
                            .setMICROESTRUCTURA4_VALOR = oPRBOND_Origen.getMICROESTRUCTURA4_VALOR
                            .setMICROESTRUCTURA5_VALOR = oPRBOND_Origen.getMICROESTRUCTURA5_VALOR
                            .setMICROESTRUCTURA6_VALOR = oPRBOND_Origen.getMICROESTRUCTURA6_VALOR
                        Else
                            .setMICROESTRUCTURA1 = ""
                            .setMICROESTRUCTURA2 = ""
                            .setMICROESTRUCTURA3 = ""
                            .setMICROESTRUCTURA4 = ""
                            .setMICROESTRUCTURA5 = ""
                            .setMICROESTRUCTURA6 = ""
                            .setMICROESTRUCTURA1_R = ""
                            .setMICROESTRUCTURA2_R = ""
                            .setMICROESTRUCTURA3_R = ""
                            .setMICROESTRUCTURA4_R = ""
                            .setMICROESTRUCTURA5_R = ""
                            .setMICROESTRUCTURA6_R = ""
                            .setMICROESTRUCTURA1_VALOR = ""
                            .setMICROESTRUCTURA2_VALOR = ""
                            .setMICROESTRUCTURA3_VALOR = ""
                            .setMICROESTRUCTURA4_VALOR = ""
                            .setMICROESTRUCTURA5_VALOR = ""
                            .setMICROESTRUCTURA6_VALOR = ""
                        End If
                        If oFichaBond.getTRACCION <> 0 Then
                            .setTRACCION = oPRBOND_Origen.getTRACCION
                            .setTRACCION_RES = oPRBOND_Origen.getTRACCION_RES
                            .setTRACCION_PASS = oPRBOND_Origen.getTRACCION_PASS
                        Else
                            .setTRACCION = ""
                            .setTRACCION_RES = ""
                            .setTRACCION_PASS = 2
                        End If
                        If oFichaBond.getMACRO_DUREZA <> 0 Then
                            .setMACRO_DUREZA = oPRBOND_Origen.getMACRO_DUREZA
                            .setMACRO_DUREZA_DIMENSION = oPRBOND_Origen.getMACRO_DUREZA_DIMENSION
                            .setMACRO_DUREZA_ESPESOR = oPRBOND_Origen.getMACRO_DUREZA_ESPESOR
                            .setMACRO_DUREZA_RES = oPRBOND_Origen.getMACRO_DUREZA_RES
                            .setMACRO_DUREZA_SD = Replace(oPRBOND_Origen.getMACRO_DUREZA_SD, ",", ".")
                            .setMACRO_DUREZA_POR = Replace(oPRBOND_Origen.getMACRO_DUREZA_POR, ",", ".")
                            .setMACRO_DUREZA_PASS = oPRBOND_Origen.getMACRO_DUREZA_PASS
                        Else
                            .setMACRO_DUREZA = ""
                            .setMACRO_DUREZA_DIMENSION = ""
                            .setMACRO_DUREZA_ESPESOR = ""
                            .setMACRO_DUREZA_RES = ""
                            .setMACRO_DUREZA_SD = 0
                            .setMACRO_DUREZA_POR = 0
                            .setMACRO_DUREZA_PASS = 2
                        End If
                        If oFichaBond.getMICRO_DUREZA <> 0 Then
                            .setMICRO_DUREZA = oPRBOND_Origen.getMICRO_DUREZA
                            .setMICRO_DUREZA_RES = oPRBOND_Origen.getMICRO_DUREZA_RES
                            .setMICRO_DUREZA_SD = Replace(oPRBOND_Origen.getMICRO_DUREZA_SD, ",", ".")
                            .setMICRO_DUREZA_POR = Replace(oPRBOND_Origen.getMICRO_DUREZA_POR, ",", ".")
                            .setMICRO_DUREZA_PASS = oPRBOND_Origen.getMICRO_DUREZA_PASS
                        Else
                            .setMICRO_DUREZA = ""
                            .setMICRO_DUREZA_RES = ""
                            .setMICRO_DUREZA_SD = 0
                            .setMICRO_DUREZA_POR = 0
                            .setMICRO_DUREZA_PASS = 2
                        End If
                        If oFichaBond.getESPESOR <> 0 Then
                            .setESPESOR = oPRBOND_Origen.getESPESOR
                            .setESPESOR_MIN = oPRBOND_Origen.getESPESOR_MIN
                            .setESPESOR_MAX = oPRBOND_Origen.getESPESOR_MAX
                            .setESPESOR_RES = oPRBOND_Origen.getESPESOR_RES
                            .setESPESOR_SD = Replace(oPRBOND_Origen.getESPESOR_SD, ",", ".")
                            .setESPESOR_POR = Replace(oPRBOND_Origen.getESPESOR_POR, ",", ".")
                            .setESPEROR_PASS = oPRBOND_Origen.getESPESOR_PASS
                        Else
                            .setESPESOR = ""
                            .setESPESOR_MIN = ""
                            .setESPESOR_MAX = ""
                            .setESPESOR_RES = ""
                            .setESPESOR_SD = 0
                            .setESPESOR_POR = 0
                            .setESPEROR_PASS = 2
                        End If
                        .Insertar
                    End If
                    If existePRTOP Then
                        ' TOP
                        .setMUESTRA_ID = ID_MUESTRA
                        .setTIPO = 2
                        .setBATCH = oPRTOP_Origen.getBATCH
                        If oFichaTop.getMICROESTRUCTURA <> 0 Then
                            .setMICROESTRUCTURA1 = oPRTOP_Origen.getMICROESTRUCTURA1
                            .setMICROESTRUCTURA2 = oPRTOP_Origen.getMICROESTRUCTURA2
                            .setMICROESTRUCTURA3 = oPRTOP_Origen.getMICROESTRUCTURA3
                            .setMICROESTRUCTURA4 = oPRTOP_Origen.getMICROESTRUCTURA4
                            .setMICROESTRUCTURA5 = oPRTOP_Origen.getMICROESTRUCTURA5
                            .setMICROESTRUCTURA6 = oPRTOP_Origen.getMICROESTRUCTURA6
                            .setMICROESTRUCTURA1_R = oPRTOP_Origen.getMICROESTRUCTURA1_R
                            .setMICROESTRUCTURA2_R = oPRTOP_Origen.getMICROESTRUCTURA2_R
                            .setMICROESTRUCTURA3_R = oPRTOP_Origen.getMICROESTRUCTURA3_R
                            .setMICROESTRUCTURA4_R = oPRTOP_Origen.getMICROESTRUCTURA4_R
                            .setMICROESTRUCTURA5_R = oPRTOP_Origen.getMICROESTRUCTURA5_R
                            .setMICROESTRUCTURA6_R = oPRTOP_Origen.getMICROESTRUCTURA6_R
                            .setMICROESTRUCTURA1_VALOR = oPRTOP_Origen.getMICROESTRUCTURA1_VALOR
                            .setMICROESTRUCTURA2_VALOR = oPRTOP_Origen.getMICROESTRUCTURA2_VALOR
                            .setMICROESTRUCTURA3_VALOR = oPRTOP_Origen.getMICROESTRUCTURA3_VALOR
                            .setMICROESTRUCTURA4_VALOR = oPRTOP_Origen.getMICROESTRUCTURA4_VALOR
                            .setMICROESTRUCTURA5_VALOR = oPRTOP_Origen.getMICROESTRUCTURA5_VALOR
                            .setMICROESTRUCTURA6_VALOR = oPRTOP_Origen.getMICROESTRUCTURA6_VALOR
                        Else
                            .setMICROESTRUCTURA1 = ""
                            .setMICROESTRUCTURA2 = ""
                            .setMICROESTRUCTURA3 = ""
                            .setMICROESTRUCTURA4 = ""
                            .setMICROESTRUCTURA5 = ""
                            .setMICROESTRUCTURA6 = ""
                            .setMICROESTRUCTURA1_R = ""
                            .setMICROESTRUCTURA2_R = ""
                            .setMICROESTRUCTURA3_R = ""
                            .setMICROESTRUCTURA4_R = ""
                            .setMICROESTRUCTURA5_R = ""
                            .setMICROESTRUCTURA6_R = ""
                            .setMICROESTRUCTURA1_VALOR = ""
                            .setMICROESTRUCTURA2_VALOR = ""
                            .setMICROESTRUCTURA3_VALOR = ""
                            .setMICROESTRUCTURA4_VALOR = ""
                            .setMICROESTRUCTURA5_VALOR = ""
                            .setMICROESTRUCTURA6_VALOR = ""
                        End If
                        If oFichaTop.getTRACCION <> 0 Then
                            .setTRACCION = oPRTOP_Origen.getTRACCION
                            .setTRACCION_RES = oPRTOP_Origen.getTRACCION_RES
                            .setTRACCION_PASS = oPRTOP_Origen.getTRACCION_PASS
                        Else
                            .setTRACCION = ""
                            .setTRACCION_RES = ""
                            .setTRACCION_PASS = 2
                        End If
                        If oFichaTop.getMACRO_DUREZA <> 0 Then
                            .setMACRO_DUREZA = oPRTOP_Origen.getMACRO_DUREZA
                            .setMACRO_DUREZA_DIMENSION = oPRTOP_Origen.getMACRO_DUREZA_DIMENSION
                            .setMACRO_DUREZA_ESPESOR = oPRTOP_Origen.getMACRO_DUREZA_ESPESOR
                            .setMACRO_DUREZA_RES = oPRTOP_Origen.getMACRO_DUREZA_RES
                            .setMACRO_DUREZA_SD = Replace(oPRTOP_Origen.getMACRO_DUREZA_SD, ",", ".")
                            .setMACRO_DUREZA_POR = Replace(oPRTOP_Origen.getMACRO_DUREZA_POR, ",", ".")
                            .setMACRO_DUREZA_PASS = oPRTOP_Origen.getMACRO_DUREZA_PASS
                        Else
                            .setMACRO_DUREZA = ""
                            .setMACRO_DUREZA_DIMENSION = ""
                            .setMACRO_DUREZA_ESPESOR = ""
                            .setMACRO_DUREZA_RES = ""
                            .setMACRO_DUREZA_SD = 0
                            .setMACRO_DUREZA_POR = 0
                            .setMACRO_DUREZA_PASS = 2
                        End If
                        If oFichaTop.getMICRO_DUREZA <> 0 Then
                            .setMICRO_DUREZA = oPRTOP_Origen.getMICRO_DUREZA
                            .setMICRO_DUREZA_RES = oPRTOP_Origen.getMICRO_DUREZA_RES
                            .setMICRO_DUREZA_SD = Replace(oPRTOP_Origen.getMICRO_DUREZA_SD, ",", ".")
                            .setMICRO_DUREZA_POR = Replace(oPRTOP_Origen.getMICRO_DUREZA_POR, ",", ".")
                            .setMICRO_DUREZA_PASS = oPRTOP_Origen.getMICRO_DUREZA_PASS
                        Else
                            .setMICRO_DUREZA = ""
                            .setMICRO_DUREZA_RES = ""
                            .setMICRO_DUREZA_SD = 0
                            .setMICRO_DUREZA_POR = 0
                            .setMICRO_DUREZA_PASS = 2
                        End If
                        If oFichaTop.getESPESOR <> 0 Then
                            .setESPESOR = oPRTOP_Origen.getESPESOR
                            .setESPESOR_MIN = oPRTOP_Origen.getESPESOR_MIN
                            .setESPESOR_MAX = oPRTOP_Origen.getESPESOR_MAX
                            .setESPESOR_RES = oPRTOP_Origen.getESPESOR_RES
                            .setESPESOR_SD = Replace(oPRTOP_Origen.getESPESOR_SD, ",", ".")
                            .setESPESOR_POR = Replace(oPRTOP_Origen.getESPESOR_POR, ",", ".")
                            .setESPEROR_PASS = oPRTOP_Origen.getESPESOR_PASS
                        Else
                            .setESPESOR = ""
                            .setESPESOR_MIN = ""
                            .setESPESOR_MAX = ""
                            .setESPESOR_RES = ""
                            .setESPESOR_SD = 0
                            .setESPESOR_POR = 0
                            .setESPEROR_PASS = 2
                        End If
                        .Insertar
                    End If
                End With
                ' PLASMA_RESULTADOS_HISTORICO
                Dim oPRH As New clsPlasma_resultados_historico
                oPRH.generar ID_MUESTRA
                Set oPRH = Nothing
                ' PLASMA_TRACCION
                'BOND
                With oPTRA
                    If existePTRABOND Then
                        .setMUESTRA_ID = ID_MUESTRA
                        .setTIPO = 1
                        .setROOM = oPTRABOND.getROOM
                        .setVELOCITY = oPTRABOND.getVELOCITY
                        .setEPOXY = oPTRABOND.getEPOXY
                        .setADHESIVE = oPTRABOND.getADHESIVE
                        .setESPESOR = oPTRABOND.getESPESOR
                        .setAVERAGE = oPTRABOND.getAVERAGE
                        .setSD = oPTRABOND.getSD
                        .Insertar
                    End If
                    If existePTRATOP Then
                        .setMUESTRA_ID = ID_MUESTRA
                        .setTIPO = 2
                        .setROOM = oPTRATOP.getROOM
                        .setVELOCITY = oPTRATOP.getVELOCITY
                        .setEPOXY = oPTRATOP.getEPOXY
                        .setADHESIVE = oPTRATOP.getADHESIVE
                        .setESPESOR = oPTRATOP.getESPESOR
                        .setAVERAGE = oPTRATOP.getAVERAGE
                        .setSD = oPTRATOP.getSD
                        .Insertar
                    End If
                End With
                ' PLASMA_TRACCION_P
                Dim oPTP As New clsPlasma_traccion_p
                Dim rsPTP As ADODB.Recordset
                oPTP.Eliminar ID_MUESTRA
                ' BOND
                Set rsPTP = oPTP.Listado(PK, 1)
                If rsPTP.RecordCount > 0 Then
                    Do
                        With oPTP
                            .setMUESTRA_ID = ID_MUESTRA
                            .setTIPO = 1
                            .setORDEN = rsPTP("ORDEN")
                            .setIDENTIFICATION = rsPTP("IDENTIFICATION")
                            .setDIAMETER = Replace(rsPTP("DIAMETER"), ",", ".")
                            .setAREA = Replace(rsPTP("AREA"), ",", ".")
                            .setLOADP = Replace(rsPTP("LOADP"), ",", ".")
                            .setTENSILE = Replace(rsPTP("TENSILE"), ",", ".")
                            .setLOCATION = rsPTP("LOCATION")
                            .Insertar
                        End With
                        rsPTP.MoveNext
                    Loop Until rsPTP.EOF
                End If
                ' TOP
                Set rsPTP = oPTP.Listado(PK, 2)
                If rsPTP.RecordCount > 0 Then
                    Do
                        With oPTP
                            .setMUESTRA_ID = ID_MUESTRA
                            .setTIPO = 2
                            .setORDEN = rsPTP("ORDEN")
                            .setIDENTIFICATION = rsPTP("IDENTIFICATION")
                            .setDIAMETER = Replace(rsPTP("DIAMETER"), ",", ".")
                            .setAREA = Replace(rsPTP("AREA"), ",", ".")
                            .setLOADP = Replace(rsPTP("LOADP"), ",", ".")
                            .setTENSILE = Replace(rsPTP("TENSILE"), ",", ".")
                            .setLOCATION = rsPTP("LOCATION")
                            .Insertar
                        End With
                        rsPTP.MoveNext
                    Loop Until rsPTP.EOF
                End If
                Set oPTP = Nothing
                Set rsPTP = Nothing
                ' ALABE
                If existePALABE Then
                    With oPAlabe
                        .setMUESTRA_ID = ID_MUESTRA
                        .Insertar
                    End With
                End If
                ' DUREZA
                If existePDUREZA Then
                    With oPDureza
                        .setMUESTRA_ID = ID_MUESTRA
                        .Insertar
                    End With
                End If
                ' PLASMA_EQUIPOS
                Dim oPEQUIPOS As New clsPlasma_equipos
                Dim rsEquipos As ADODB.Recordset
                Set rsEquipos = oPEQUIPOS.Listado(PK)
                oPEQUIPOS.Eliminar ID_MUESTRA
                If rsEquipos.RecordCount > 0 Then
                    Do
                        With oPEQUIPOS
                            .setMUESTRA_ID = ID_MUESTRA
                            .setORDEN = rsEquipos("ORDEN")
                            .setEQUIPO_ID = rsEquipos("EQUIPO_ID")
                            .setVERIFICACION_ID = rsEquipos("VERIFICACION_ID")
                            .setEN_INFORME = rsEquipos("EN_INFORME")
                            .Insertar
                        End With
                        rsEquipos.MoveNext
                    Loop Until rsEquipos.EOF
                End If
                Set oPEQUIPOS = Nothing
                Set rsEquipos = Nothing
                ' PLASMA_REACTIVOS
                Dim oPREACTIVOS As New clsPlasma_Reactivos
                Dim rsReactivos As ADODB.Recordset
                Set rsReactivos = oPREACTIVOS.Listado(PK)
                oPREACTIVOS.Eliminar ID_MUESTRA
                If rsReactivos.RecordCount > 0 Then
                    Do
                        With oPREACTIVOS
                            .setMUESTRA_ID = ID_MUESTRA
                            .setBOTE_EX_ID = rsReactivos("BOTE_EX_ID")
                            .setTIPO = rsReactivos("TIPO")
                            .setORDEN = rsReactivos("ORDEN")
                            .Insertar
                        End With
                        rsReactivos.MoveNext
                    Loop Until rsReactivos.EOF
                End If
                Set oPREACTIVOS = Nothing
                Set rsReactivos = Nothing
                ' IMAGENES
                Dim oDocumentacion As New clsDocumentacion
                oDocumentacion.MuestrasImagenesDuplicar PK, ID_MUESTRA
                Set oDocumentacion = Nothing
            End If
        Next
        MsgBox "Resultados duplicados correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPlasma_CopiaResultados"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera_lista
    cargar_lista
End Sub
Private Sub cabecera_lista()
    With lista.ColumnHeaders
        .Add , , "Código", 1600, lvwColumnLeft
        .Add , , "Cliente", 2000, lvwColumnLeft
        .Add , , "Tipo de Analisis/Solución", 1200, lvwColumnLeft
        .Add , , "Ref.Cliente", 3000, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "General", 800, lvwColumnCenter
        .Add , , "ID_MUESTRA", 1, lvwColumnCenter
        .Add , , "Facturada", 1, lvwColumnCenter
        .Add , , "Centro", 1200, lvwColumnCenter
        .Add , , "", 250, lvwColumnCenter
    End With
End Sub
Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim consulta As String
   On Error GoTo cargar_lista_Error

    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada,mu.revision_usuario,ce.nombre,mu.situacion " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "centros as ce, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND mu.tipo_muestra_id=tm.id_tipo_muestra AND mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                      " and mu.centro_id = ce.id_centro " & _
                      " and mu.anulada = 0 " & _
                      " and mu.cerrada = 0 " & _
                      " and mu.analisis_modificado = " & tipo_especial.PLASMA & _
                      " and mu.tipo_muestra_id = " & TIPOS_MUESTRAS.PLASMA & _
                      " order by mu.id_general desc"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    lista.ListItems.Clear
    Dim i As Integer
    If rs.RecordCount >= 1 Then
        i = 1
        Dim objLitem As ListItem, objSI As ListSubItem
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(1))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(8)
                .SubItems(3) = rs.Fields(4)
                If Not IsNull(rs.Fields(5)) Then
                .SubItems(4) = rs.Fields(5)
                End If
                If Not IsNull(rs.Fields(9)) Then
                   .SubItems(5) = Format(rs.Fields(9), "00000")
                End If
                If Not IsNull(rs.Fields(6)) Then
                    .SubItems(6) = rs.Fields(6)
                End If
                .SubItems(7) = rs(10)
                .SubItems(8) = rs(15) 'CENTRO
                    
                If rs(13) = 1 Then ' Si cerrada, bola de color
                    .ListSubItems.Add , , "", rs(16) + 7
                End If
           '     .SubItems(9) = rs(16) 'SITUACION
            End With
            i = lista.ListItems.Count
            lista.ListItems(i).Checked = False
            If rs.Fields(11) <> 0 Then 'ENVIADO_CORREO
                lista.ListItems(i).SmallIcon = 1
                lista.ListItems(i).ToolTipText = "Enviado Correo"
            Else
                If rs(12) <> 0 Then ' ANULADA
                    lista.ListItems(i).SmallIcon = 2
                    lista.ListItems(i).ToolTipText = "Anulada"
                Else
                    Select Case rs(13) ' Cerrada
                        Case 0 ' Abierta
                            lista.ListItems(i).SmallIcon = 5
                            lista.ListItems(i).ToolTipText = "Abierta"
                        Case 1 ' Cerrada
                            If rs(14) = 0 Then ' Revision Usuario
                                lista.ListItems(i).SmallIcon = 6
                                lista.ListItems(i).ToolTipText = "Cerrada Pendiente Revisar"
                            Else
                                lista.ListItems(i).SmallIcon = 4
                                lista.ListItems(i).ToolTipText = "Cerrada y Revisada por Usuario : " & rs(14)
                            End If
                        Case 2 ' Pdte. Cierre
                            lista.ListItems(i).SmallIcon = 3
                            lista.ListItems(i).ToolTipText = "Pdte. Cierre"
                    End Select
                End If
            End If
            rs.MoveNext
        Wend
    End If
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmPlasma_CopiaResultados"

End Sub
Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub
