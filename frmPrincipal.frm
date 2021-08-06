VERSION 5.00
Begin VB.Form frmPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Ficha de Consulta Prévia (2005)"
   ClientHeight    =   9180
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11685
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Resposta"
      ForeColor       =   &H00800000&
      Height          =   3600
      Left            =   30
      TabIndex        =   12
      Top             =   5490
      Width           =   11445
      Begin VB.ListBox List3 
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00008000&
         Height          =   2010
         ItemData        =   "frmPrincipal.frx":000C
         Left            =   90
         List            =   "frmPrincipal.frx":000E
         TabIndex        =   13
         Top             =   270
         Width           =   11175
      End
      Begin VB.Label Label8 
         Caption         =   "Ateção as leis apresentadas aqui são referentes ao ano 2005. Esse projeto é apenas um exemplo de programação."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1065
         Left            =   165
         TabIndex        =   19
         Top             =   2370
         Width           =   11055
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Local referente a sua consulta"
      ForeColor       =   &H00800000&
      Height          =   2895
      Left            =   30
      TabIndex        =   7
      Top             =   2520
      Width           =   11445
      Begin VB.ListBox List2 
         ForeColor       =   &H00008000&
         Height          =   1035
         ItemData        =   "frmPrincipal.frx":0010
         Left            =   90
         List            =   "frmPrincipal.frx":0012
         TabIndex        =   9
         Top             =   1080
         Width           =   8655
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   450
         Width           =   8655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trecho escolhido"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Label LabelTrecho 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   2340
         Width           =   7305
      End
      Begin VB.Label LabelZona 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7470
         TabIndex        =   16
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zoneamento"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7470
         TabIndex        =   15
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[Exemplo ]"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   8820
         TabIndex        =   14
         Top             =   270
         Width           =   735
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2190
         Left            =   8820
         Top             =   450
         Width           =   2520
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Logradouro"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Escolha o Trecho do Logradouro referente ao local da consulta"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   900
         Width           =   4485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Escolha a(s) atividade(s) pretendida que deseja estabelecer na cidade"
      ForeColor       =   &H00800000&
      Height          =   2265
      Left            =   30
      TabIndex        =   3
      Top             =   180
      Width           =   11445
      Begin VB.TextBox utilizada 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9540
         TabIndex        =   2
         Top             =   1080
         Width           =   1760
      End
      Begin VB.TextBox Total 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9540
         TabIndex        =   1
         Top             =   360
         Width           =   1760
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   1905
         Left            =   6850
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ListBox List1 
         ForeColor       =   &H00008000&
         Height          =   1860
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   180
         Width           =   9330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Área a ser utilizada (m2)"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9540
         TabIndex        =   6
         Top             =   900
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área total do imóvel (m2)"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9540
         TabIndex        =   5
         Top             =   180
         Width           =   1740
      End
   End
   Begin VB.Menu menarq 
      Caption         =   "Arquivo"
      Begin VB.Menu mensair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu menajuda 
      Caption         =   "Ajuda"
      Begin VB.Menu menhelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mensobre 
         Caption         =   "Sobre"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
        Dim N As String
        Dim Z, W As Integer
        Dim existe As Boolean
        existe = False
        
        'LENDO LOGRADOUROS
        Me.MousePointer = 11
        For Z = 1 To 3273
        Get #5, Z, DLOGRA
        If Trim(Combo1.Text) = Trim(DLOGRA.B) Then
           N = Format(DLOGRA.A, "0000")
           Exit For
        End If
        Next Z
       
       'LENDO TRECHOS
        List2.Clear
        On Error Resume Next
        Open App.Path & "\trechos.dat" For Random As #6 Len = Len(DTRE)
        For W = 1 To 1233
            Get #6, W, DTRE
            If Trim(DTRE.B) = Trim(N) Then
               existe = True
               If (Trim(DTRE.D) <> "" And Trim(DTRE.E) <> "") Then 'TEM TRECHO
                  List2.AddItem DTRE.A & " - " & Trim(NOMELOG(DTRE.D)) & "  /  " & Trim(NOMELOG(DTRE.E))
               Else
                  List2.AddItem DTRE.A & " - " & Trim(NOMELOG(N))
               End If
            End If
        Next W
        
        If existe = False Then
           MsgBox "Logradouro não existe no sistema,aguarde, em breve estaremos disponibilizando", vbInformation, "Prezado usuário"
           Combo1.SetFocus
        End If
        Me.MousePointer = 0


End Sub

Private Sub Form_Load()
    
    Dim W, Z As Integer
   'LENDO ATIVIDADES
    Me.MousePointer = 11
    List1.Clear
    On Error Resume Next
    Open App.Path & "\atividades.dat" For Random As #4 Len = Len(DATI)
    For W = 1 To 840
        Get #4, W, DATI
        List1.AddItem DATI.B
    Next W
    Me.MousePointer = 0
    Close #4
    
    'LENDO LOGRADOUROS
    Me.MousePointer = 11
    Combo1.Clear
    On Error Resume Next
    Open App.Path & "\Logradouros.dat" For Random As #5 Len = Len(DLOGRA)
    For Z = 1 To 3273
        Get #5, Z, DLOGRA
        Combo1.AddItem DLOGRA.B
    Next Z
    Me.MousePointer = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Visible = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &H800000
End Sub

Private Sub Image1_Click()
           MsgBox "Aguarde, sistema em desenvolvimento", vbInformation, "Prezado usuário"
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &HFFFF&
End Sub

Private Sub LabelTrecho_Change()

        On Error Resume Next
        Open App.Path & "\trechos.dat" For Random As #6 Len = Len(DTRE)
            Get #6, Val(LabelTrecho), DTRE
            LabelZona.Caption = DTRE.C
            VZONA = LabelZona.Caption
        Close #6
        
End Sub

Private Sub List2_Click()

    VTRECHO = Trim(List2.List(List2.ListIndex))
    LabelTrecho.Caption = List2.List(List2.ListIndex)
    Image1.Picture = LoadPicture(App.Path & "\dexemplo.gif")
    Call RESPONDER
    
End Sub

Private Sub menhelp_Click()
       MsgBox "Aguarde, Help em desenvolvimento", vbInformation, "Prezado usuário"
End Sub

Private Sub mensair_Click()
        Close
        MsgBox "A Prefeitura de Uberaba lhe deseja sucesso no seu empreendimento!", vbInformation, "Prezado usuário"
        Unload Me
        End
End Sub

Private Sub mensobre_Click()
        frmAbout.Show 1
End Sub
Private Sub List1_Click()
    If Len(Trim(List1.List(List1.ListIndex))) > 110 Then
        Text1.Visible = True
        Text1.Text = ""
        Text1.Text = List1.List(List1.ListIndex)
    Else
        Text1.Text = ""
        Text1.Visible = False
    End If
End Sub

Private Sub List1_ItemCheck(Item As Integer)
    If Len(Trim(List1.List(List1.ListIndex))) > 110 Then
        Text1.Visible = True
        Text1.Text = ""
        Text1.Text = List1.List(List1.ListIndex)
    Else
        Text1.Text = ""
        Text1.Visible = False
    End If
    
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(Trim(List1.List(List1.ListIndex))) > 110 Then
        Text1.Visible = True
        Text1.Text = ""
        Text1.Text = List1.List(List1.ListIndex)
    Else
        Text1.Text = ""
        Text1.Visible = False
    End If
End Sub
Public Function NOMELOG(S As String) As String
        Dim C As Integer
        'LENDO LOGRADOUROS
       For C = 1 To 3273
         Get #5, C, DLOGRA
         If DLOGRA.A = S Then
           NOMELOG = DLOGRA.B
           Exit For
         End If
       Next C
 
End Function


Public Sub RESPONDER()
    Dim A As Integer
    A = 0
    
   'VALIDAÇÃO DE ATIVIDADE
    If List1.SelCount = 0 Then
       MsgBox "É necessário Selecionar a(s) atividade(s)", vbInformation, "Prezado usuário"
       F1.List1.SetFocus
       Exit Sub
    End If

   'VALIDA AREA TOTAL
   If Val(Total.Text) = 0 Then
       MsgBox "É necessário informar o tamanho total da área do local da consulta!", vbInformation, "Prezado usuário"
       Total.SetFocus
       Exit Sub
   Else
       VTOTAL = Val(Total.Text)
   End If

   'VALIDA AREA UTILIZADA
   If Val(utilizada.Text) = 0 Then
       MsgBox "É necessário informar a área a ser utilizada pela atividade!", vbInformation, "Prezado usuário"
       utilizada.SetFocus
       Exit Sub
   Else
        VUTILIZADA = Val(utilizada.Text)
   End If
   
   'VALIDA AREA UTILIZADA
   If VUTILIZADA > VTOTAL Then
       MsgBox "Atenção, a área a ser utilizada não pode ser maior que o total da área!", vbInformation, "Prezado usuário"
       utilizada.SetFocus
       Exit Sub
   End If
   
    Me.List3.Clear
    For k = 0 To F1.List1.ListCount - 1        'PARA CADA ATIVIDADE
         If F1.List1.Selected(k) = True Then   'SE SELECIONADA NA LISTA
           'TRATA VARIAVEIS
           VNOMEOBS = ""
           VLEICOMP = ""
           VPERMISSAO = ""
           VRESULTADO = ""
           A = A + 1
           
           VNTIPO = BUSCANTIPO(k + 1)
           VNOMETIPO = BUSCANOMETIPO(Trim(Str(VNTIPO)))
           VNUSO = BUSCANUSO(Trim(Str(VNTIPO)))
           VNOMEUSO = BUSCANOMEUSO(Trim(Str(VNUSO)))
           VNOMEPORTE = BUSCANOMEPORTE(Trim(Str(VNUSO)))
           Call BUSCALEI(F1.List1.ListIndex + 1)
           If Trim(VLETRA) = "" Then
                  VNOMEPORTE = "INVÁLIDO!"
                  VPERMISSAO = "PORTE INVÁLIDO PARA ESTA ATIVIDADE!"
           Else
                  Call BUSCAPERMISSAO
           End If
           
           Me.List3.AddItem "ANÁLISE= " & A & "  USO: " & VNOMEUSO & "   TIPO: " & VNOMETIPO
           Me.List3.AddItem "ATIVIDADE: " & UCase((Me.List1.List(k)))
           Me.List3.AddItem "LOGRADOURO: " & Trim(Me.Combo1.Text) & "   " & "TRECHO: " & VTRECHO
           Me.List3.AddItem "PORTE:  " & VUTILIZADA & "m2" & " =  " & VNOMEPORTE & "       ZONEAMENTO: " & VZONA
           Me.List3.AddItem "OBS:  " & VNOMEOBS
           Me.List3.AddItem "LEI COMPLEMENTAR:  " & VLEICOMP
           If VNOMEPORTE = "INVÁLIDO!" Then
              VRESULTADO = "PORTE INVÁLIDO PARA ESTA ATIVIDADE."
           End If
           Me.List3.AddItem "RESULTADO:  " & VRESULTADO
           Me.List3.AddItem ""
        End If
    Next k
    
End Sub
Public Function BUSCANTIPO(k As Integer) As Integer
       'ABRIR AS ATIVIDADES PEGAR CODIGO DO TIPO
       On Error Resume Next
       Open App.Path & "\atividades.dat" For Random As #4 Len = Len(DATI)
       Get #4, k, DATI
       BUSCANTIPO = DATI.C
       Close #4
End Function
Public Sub BUSCALEI(N As Integer)
       Dim J As Integer
       VLETRA = ""
       VOBS = ""
       'LENDO LEI
       On Error Resume Next
       Open App.Path & "\lei.dat" For Random As #3 Len = Len(DLEI)
       Get #3, N, DLEI
       Select Case VPORTE
       Case 1
            VLETRA = DLEI.B
            VOBS = DLEI.C
       Case 2
            VLETRA = DLEI.D
            VOBS = DLEI.E
       Case 3
            VLETRA = DLEI.F
            VOBS = DLEI.G
       End Select
       VLEICOMP = DLEI.J
       Close #3
       
       If (Trim(VOBS) <> "") Then
         'LENDO OBSERVACOES
         On Error Resume Next
         Open App.Path & "\observacoes.dat" For Random As #8 Len = Len(DOBS)
         Get #8, Val(Trim(VOBS)), DOBS
         VNOMEOBS = Trim(UCase(DOBS.B))
         Close #8
       End If
       
End Sub
Public Function BUSCANOMETIPO(T As String) As String
       'LENDO TIPOS
       On Error Resume Next
       Open App.Path & "\tipos.dat" For Random As #2 Len = Len(DTIPO)
       Get #2, Val(T), DTIPO
       BUSCANOMETIPO = DTIPO.B
       Close #2
End Function
Public Function BUSCANUSO(U As String) As String
       'LENDO TIPOS
       On Error Resume Next
       Open App.Path & "\tipos.dat" For Random As #2 Len = Len(DTIPO)
       Get #4, Val(U), DTIPO
       BUSCANUSO = DTIPO.C
       Close #4
End Function
Public Function BUSCANOMEUSO(U As String) As String
       'LENDO USOS
       On Error Resume Next
       Open App.Path & "\usos.dat" For Random As #1 Len = Len(DUSO)
       Get #1, Val(U), DUSO
       BUSCANOMEUSO = DUSO.B
       Close #1
End Function
Public Function BUSCANOMEPORTE(numero As String) As String
       'LENDO USOS
       On Error Resume Next
       Open App.Path & "\usos.dat" For Random As #1 Len = Len(DUSO)
       Get #1, Val(numero), DUSO
       If (VUTILIZADA > DUSO.C) And (VUTILIZADA <= DUSO.D) Then
          BUSCANOMEPORTE = "PP"
          VPORTE = 1
       End If
       If (VUTILIZADA > DUSO.D) And (VUTILIZADA <= DUSO.F) Then
          BUSCANOMEPORTE = "MP"
          VPORTE = 2
       End If
       If (VUTILIZADA > DUSO.F) Then
          BUSCANOMEPORTE = "GP"
          VPORTE = 3
       End If
       Close #1
End Function
Public Sub BUSCAPERMISSAO()
       Open App.Path & "\permissoes.dat" For Random As #9 Len = Len(DPERMI)
       For J = 1 To 13
            Get #9, J, DPERMI
            If Trim(VZONA) = Trim(DPERMI.B) Then
                    If Trim(VLETRA) = "A" Then
                       VPERMISSAO = DPERMI.C
                    End If
                    If Trim(VLETRA) = "B" Then
                        VPERMISSAO = DPERMI.D
                    End If
                    If Trim(VLETRA) = "C" Then
                        VPERMISSAO = DPERMI.E
                    End If
                    If Trim(VLETRA) = "D" Then
                        VPERMISSAO = DPERMI.F
                    End If
                    If Trim(VLETRA) = "E" Then
                        VPERMISSAO = DPERMI.G
                    End If
                    If Trim(VLETRA) = "F" Then
                        VPERMISSAO = DPERMI.H
                    End If
                    If Trim(VLETRA) = "G" Then
                        VPERMISSAO = DPERMI.I
                    End If
            End If
            
        Next J
        Close #9
        
        If VPERMISSAO = "S" Then
            VRESULTADO = "APROVADA!"
        Else
            VRESULTADO = "VETADA!"
        End If
        
End Sub


