Attribute VB_Name = "Cavalete"
Option Explicit

' =========================================================
' FACHADA DE COMPATIBILIDADE
' Mantém as macros públicas originais e delega para o módulo core.
' =========================================================

' Insere o cavalete cinza no documento.
Public Sub CavaleteCinza()
    InserirCavalete CAMINHO_CAVALETE_CZ, NOME_GRUPO_CZ
End Sub

' Insere o cavalete branco no documento.
Public Sub CavaleteBranco()
    InserirCavalete CAMINHO_CAVALETE_BR, NOME_GRUPO_BR
End Sub

' Insere o cavalete preto no documento.
Public Sub CavaletePreto()
    InserirCavalete CAMINHO_CAVALETE_PT, NOME_GRUPO_PT
End Sub


