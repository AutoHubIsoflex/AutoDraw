Attribute VB_Name = "Core"
Option Explicit

' ==============================================================================
' CORE - APLICAÇÃO DAS MOLDURAS
' ==============================================================================

Public Sub AplicarMolduraPadrao(ByVal caminhoArquivoMoldura As String)

    On Error GoTo TrataErro

    Dim retanguloBase As Shape
    Dim grupoMoldura As Shape
    Dim cantSupDir As Shape, cantSupEsq As Shape
    Dim cantInfEsq As Shape, cantInfDir As Shape
    Dim tuboDir As Shape, tuboSup As Shape
    Dim tuboEsq As Shape, tuboInf As Shape
    Dim deslocamentoMoldura As Double

    deslocamentoMoldura = MmParaUnidadeDocumento(OFFSET_MOLDURA_PADRAO_MM)

    If Not TentarObterRetanguloBaseMagenta(retanguloBase) Then Exit Sub
    If Not ImportarMolduraEObterGrupo(caminhoArquivoMoldura, "Arquivo não encontrado: ", True, grupoMoldura) Then Exit Sub

    MapearPecasBasicas grupoMoldura, cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
                       tuboDir, tuboSup, tuboEsq, tuboInf

    If Not ValidarCantoneiras(cantSupDir, cantSupEsq, cantInfEsq, cantInfDir) Then Exit Sub
    If Not ValidarTubos(tuboDir, tuboSup, tuboEsq, tuboInf) Then Exit Sub

    PosicionarCantoneiras retanguloBase, deslocamentoMoldura, _
                          cantSupDir, cantSupEsq, cantInfEsq, cantInfDir

    PosicionarTubosBasicos retanguloBase, cantSupDir, cantSupEsq, cantInfDir, _
                           tuboDir, tuboSup, tuboEsq, tuboInf

    AjustarTubosEntreCentros cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
                             tuboDir, tuboSup, tuboEsq, tuboInf

    Exit Sub

TrataErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "AplicarMolduraPadrao"

End Sub

Public Sub AplicarMolduraEconomy(ByVal caminhoArquivoMoldura As String)

    On Error GoTo TrataErro

    Dim retanguloBase As Shape
    Dim grupoMoldura As Shape

    Dim cantSupDir As Shape, cantSupEsq As Shape
    Dim cantInfEsq As Shape, cantInfDir As Shape

    Dim tuboDir As Shape, tuboSup As Shape
    Dim tuboEsq As Shape, tuboInf As Shape

    Dim alhetaInfDir As Shape, alhetaInfEsq As Shape
    Dim alhetaSupDir As Shape, alhetaSupEsq As Shape

    Dim alhetaInfEsqDup As Shape, alhetaSupEsqDup As Shape
    Dim alhetaInfEsqDupExtra As Shape, alhetaInfDirDupExtra As Shape
    Dim alhetaSupEsqDupExtra As Shape, alhetaSupDirDupExtra As Shape

    Dim refTuboDir As Shape, refTuboSup As Shape
    Dim refTuboEsq As Shape, refTuboInf As Shape

    Dim deslocamentoMoldura As Double
    Dim deslocamentoHorizontalAlheta As Double
    Dim larguraPeca As Double

    deslocamentoMoldura = MmParaUnidadeDocumento(DESLOCAMENTO_MOLDURA_ECONOMY_MM)
    deslocamentoHorizontalAlheta = MmParaUnidadeDocumento(DESLOCAMENTO_ALHETA_ECONOMY_MM)

    If Not TentarObterRetanguloBaseMagenta(retanguloBase) Then Exit Sub
    If Not ImportarMolduraEObterGrupo(caminhoArquivoMoldura, "Arquivo não encontrado!", False, grupoMoldura) Then Exit Sub

    MapearPecasEconomy grupoMoldura, _
        cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
        tuboDir, tuboSup, tuboEsq, tuboInf, _
        alhetaInfDir, alhetaInfEsq, alhetaSupDir, alhetaSupEsq

    If Not ValidarCantoneiras(cantSupDir, cantSupEsq, cantInfEsq, cantInfDir) Then Exit Sub
    If Not ValidarTubos(tuboDir, tuboSup, tuboEsq, tuboInf) Then Exit Sub

    If Not ValidarAlhetasInferiores(alhetaInfEsq, alhetaInfDir) Then
        MsgBox "Alhetas inferiores não encontradas!", vbCritical
        Exit Sub
    End If

    Set refTuboDir = BuscarShapePorNome(cantSupDir, NOME_REF_TUBO_DIR)
    Set refTuboSup = BuscarShapePorNome(cantSupDir, NOME_REF_TUBO_SUP)
    Set refTuboEsq = BuscarShapePorNome(cantInfEsq, NOME_REF_TUBO_ESQ)
    Set refTuboInf = BuscarShapePorNome(cantInfEsq, NOME_REF_TUBO_INF)

    PosicionarCantoneirasEconomy retanguloBase, deslocamentoMoldura, _
                                 cantSupDir, cantSupEsq, cantInfEsq, cantInfDir

    PosicionarTubosEconomy retanguloBase, _
                           cantSupDir, cantSupEsq, cantInfDir, _
                           refTuboDir, refTuboSup, refTuboEsq, refTuboInf, _
                           tuboDir, tuboSup, tuboEsq, tuboInf

    AjustarTubosEntreCentros cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
                             tuboDir, tuboSup, tuboEsq, tuboInf

    PosicionarAlhetas tuboSup, tuboInf, deslocamentoHorizontalAlheta, _
                      alhetaInfDir, alhetaInfEsq, alhetaSupDir, alhetaSupEsq

    larguraPeca = retanguloBase.SizeWidth

    If larguraPeca >= MmParaUnidadeDocumento(LARGURA_MINIMA_DUPLICAR_ALHETA_MM) _
    And larguraPeca < MmParaUnidadeDocumento(LARGURA_MINIMA_DUPLICAR_ALHETA_EXTRA_MM) Then

        Set alhetaInfEsqDup = alhetaInfEsq.Duplicate
        Set alhetaSupEsqDup = alhetaSupEsq.Duplicate

        alhetaInfEsqDup.CenterX = tuboInf.CenterX
        alhetaSupEsqDup.CenterX = tuboSup.CenterX

    End If

    If larguraPeca >= MmParaUnidadeDocumento(LARGURA_MINIMA_DUPLICAR_ALHETA_EXTRA_MM) Then

        Set alhetaInfEsqDupExtra = alhetaInfEsq.Duplicate
        Set alhetaInfDirDupExtra = alhetaInfDir.Duplicate
        Set alhetaSupEsqDupExtra = alhetaSupEsq.Duplicate
        Set alhetaSupDirDupExtra = alhetaSupDir.Duplicate

        alhetaInfEsqDupExtra.CenterX = (tuboInf.CenterX + alhetaInfEsq.CenterX) / 2
        alhetaInfDirDupExtra.CenterX = (tuboInf.CenterX + alhetaInfDir.CenterX) / 2

        alhetaSupEsqDupExtra.CenterX = (tuboSup.CenterX + alhetaSupEsq.CenterX) / 2
        alhetaSupDirDupExtra.CenterX = (tuboSup.CenterX + alhetaSupDir.CenterX) / 2

    End If

    Exit Sub

TrataErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "AplicarMolduraEconomy"

End Sub


