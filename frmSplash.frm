VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3285
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4680
      Top             =   1440
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4680
      TabIndex        =   2
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label Label2 
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
      TabIndex        =   1
      Top             =   120
      Width           =   3900
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seja bem-vindo!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   2070
   End
   Begin VB.Image imgLogo 
      Height          =   2295
      Left            =   2160
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2355
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   3255
      Left            =   0
      Picture         =   "frmSplash.frx":51C0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1380
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    T = T + 1
    If T > 7 Then
        Timer1.Enabled = False
        frmPrincipal.Show
        Unload Me
    End If
End Sub

