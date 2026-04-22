Attribute VB_Name = "Cst"
Option Explicit

' ==============================================================================
' CONSTANTES - MOLDURA AUTO
' ==============================================================================

' Caminhos dos arquivos de moldura.
Public Const PASTA_MOLDURA_AUTO As String = _
    "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\MOLDURA AUTO\"

Public Const ARQUIVO_MOLDURA_AZUL As String = PASTA_MOLDURA_AUTO & "molduraAzul.cdr"
Public Const ARQUIVO_MOLDURA_CINZA As String = PASTA_MOLDURA_AUTO & "molduraCinza.cdr"
Public Const ARQUIVO_MOLDURA_PRETO As String = PASTA_MOLDURA_AUTO & "molduraPreto.cdr"
Public Const ARQUIVO_MOLDURA_ECONOMY As String = PASTA_MOLDURA_AUTO & "molduraEconomy.cdr"

' Nomes de shapes esperados.
Public Const NOME_CANT_SUP_DIR As String = "cantSupDir"
Public Const NOME_CANT_SUP_ESQ As String = "cantSupEsq"
Public Const NOME_CANT_INF_ESQ As String = "cantInfEsq"
Public Const NOME_CANT_INF_DIR As String = "cantInfDir"
Public Const NOME_TUBO_DIR As String = "tuboDir"
Public Const NOME_TUBO_SUP As String = "tuboSup"
Public Const NOME_TUBO_ESQ As String = "tuboEsq"
Public Const NOME_TUBO_INF As String = "tuboInf"
Public Const NOME_ALHETA_INF_DIR As String = "alhetaInfDir"
Public Const NOME_ALHETA_INF_ESQ As String = "alhetaInfEsq"
Public Const NOME_ALHETA_SUP_DIR As String = "alhetaSupDir"
Public Const NOME_ALHETA_SUP_ESQ As String = "alhetaSupEsq"

Public Const NOME_REF_TUBO_DIR As String = "alinhaTuboDir"
Public Const NOME_REF_TUBO_SUP As String = "alinhaTuboSup"
Public Const NOME_REF_TUBO_ESQ As String = "alinhaTuboEsq"
Public Const NOME_REF_TUBO_INF As String = "alinhaTuboInf"

' Constantes de posicionamento (em milímetros).
Public Const OFFSET_MOLDURA_PADRAO_MM As Double = 5.46
Public Const DESLOCAMENTO_MOLDURA_ECONOMY_MM As Double = 6
Public Const DESLOCAMENTO_ALHETA_ECONOMY_MM As Double = 55
Public Const LARGURA_MINIMA_DUPLICAR_ALHETA_MM As Double = 1830
Public Const LARGURA_MINIMA_DUPLICAR_ALHETA_EXTRA_MM As Double = 2000

' Mantém uma única conversão mm -> unidade do documento.
Public Function MmParaUnidadeDocumento(ByVal valorEmMm As Double) As Double
    MmParaUnidadeDocumento = ActiveDocument.ToUnits(valorEmMm, cdrMillimeter)
End Function


