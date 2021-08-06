VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre Ficha de Consulta Prévia"
   ClientHeight    =   3075
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5475
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2122.419
   ScaleMode       =   0  'User
   ScaleWidth      =   5141.308
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   420
      Top             =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Versão 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   4800
      TabIndex        =   1
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Consulta Prévia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3900
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   0
      Picture         =   "frmAbout.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1380
   End
   Begin VB.Image imgLogo 
      Height          =   2055
      Left            =   2160
      Picture         =   "frmAbout.frx":A785
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1965
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim T As Integer


Private Sub Form_KeyPress(KeyAscii As Integer)
 Unload Me
End Sub

Private Sub Form_Load()
 Timer1.Enabled = True
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub imgLogo_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
T = T + 1
    If T > 7 Then
        Timer1.Enabled = False
        'frmPrincipal.Show
        Unload Me
    End If
End Sub
