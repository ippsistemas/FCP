Attribute VB_Name = "Module1"
Public NOME As String
Public N As Integer
Public LS, CS As String
Public LI, CI As Integer
Public VNTIPO, VNUSO, VTOTAL, VUTILIZADA, VPORTE As Integer
Public VNOMEUSO, VNOMETIPO, VNOMEPORTE, VTRECHO As String
Public VZONA, P, VPERMISSAO, VOBS, VNOMEOBS, VLETRA, VLEICOMP, VRESULTADO As String


'exporta portes
Type isaUsos
   A As String * 4
   B As String * 10
   C As Integer
   D As Integer
   E As Integer
   F As Integer
   G As Integer
   H As Integer
End Type
Public DUSO As isaUsos

'exporta tipos
Type isatipos
   A As String * 4
   B As String * 100
   C As String * 4
End Type
Public DTIPO As isatipos

'exportar lei
Type isalei
   A As String * 4
   B As String * 1
   C As String * 2
   D As String * 1
   E As String * 2
   F As String * 1
   G As String * 2
   H As String * 1
   I As String * 2
   J As String * 40
End Type
Public DLEI As isalei


'exportar atividades
Type isaatividades
   A As String * 4
   B As String * 250
   C As String * 4
End Type
Public DATI As isaatividades


'exporta logradouros
Type isalogradouros
   A As String * 4
   B As String * 80
End Type
Public DLOGRA As isalogradouros

'exporta trechos
Type isatrechos
   A As String * 4
   B As String * 4
   C As String * 8
   D As String * 4
   E As String * 4
   F As String * 4
End Type
Public DTRE As isatrechos

'exporta nomesantigos
Type isaantigos
   A As String * 4
   B As String * 80
End Type
Public DANTIGO As isaantigos

'exporta observacoes
Type isaobs
   A As String * 4
   B As String * 250
End Type
Public DOBS As isaobs

'exportar permissoes
Type isapermissoes
   A As String * 4
   B As String * 8
   C As String * 1
   D As String * 1
   E As String * 1
   F As String * 1
   G As String * 1
   H As String * 1
   I As String * 1
   J As String * 4
End Type
Public DPERMI As isapermissoes

