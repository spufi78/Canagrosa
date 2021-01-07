VERSION 5.00
Begin VB.Form frmInforme 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de fórmulas"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   Icon            =   "frmInforme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8565
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   90
      Width           =   10905
   End
End
Attribute VB_Name = "frmInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    log (Me.Name)
    Me.MousePointer = 11
    Dim c As String
    Dim rs As New ADODB.Recordset
    Dim rs_c As New ADODB.Recordset
    c = "select * From formulas order by nombre"
    Set rs = datos_bd(c)
    Do
        t = t + "--------------------------------------------------------------------------------" & vbNewLine
        t = t + "Formula : " & rs("nombre") & vbNewLine
        c = "select * From formulas_campos where formula_id=" & rs("id_formula") & " order by id_campo"
        Set rs_c = datos_bd(c)
        If rs_c.RecordCount <> 0 Then
            t = t + "Campos:" & vbNewLine
            Do
                t = t + rs_c("nombre") & "(" & rs_c("ID_CAMPO") & ")" & " Enteros:" & rs_c("enteros") & " Decimales:" & rs_c("decimales") & vbNewLine
                rs_c.MoveNext
            Loop Until rs_c.EOF
        End If
        t = t + "Campo de resultado : " & rs("campo_id_resultado") & vbNewLine
        t = t + "Expresión : " & rs("expresion") & vbNewLine
        rs.MoveNext
    Loop Until rs.EOF
    Me.MousePointer = 0
End Sub
