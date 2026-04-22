Attribute VB_Name = "Mold"
Option Explicit

' ==============================================================================
' FACHADA DE COMPATIBILIDADE
' Mantém os nomes de macro originais enquanto a lógica fica distribuída em
' módulos especializados (constantes, core, layout, mapeamento e validação).
' ==============================================================================

Public Sub molduraAzul()
    AplicarMolduraPadrao ARQUIVO_MOLDURA_AZUL
End Sub

Public Sub molduraCinza()
    AplicarMolduraPadrao ARQUIVO_MOLDURA_CINZA
End Sub

Public Sub molduraPreto()
    AplicarMolduraPadrao ARQUIVO_MOLDURA_PRETO
End Sub

Public Sub molduraEconomy()
    AplicarMolduraEconomy ARQUIVO_MOLDURA_ECONOMY
End Sub

