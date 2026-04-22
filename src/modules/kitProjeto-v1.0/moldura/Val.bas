Attribute VB_Name = "Val"
Option Explicit

' ==============================================================================
' VALIDAÇÕES E DETECÇÃO DE CONTEXTO
' ==============================================================================

Public Function ImportarMolduraEObterGrupo( _
    ByVal caminhoArquivo As String, _
    ByVal mensagemErro As String, _
    ByVal incluirCaminhoNoErro As Boolean, _
    ByRef grupoImportado As Shape) As Boolean

    If Not ValidarArquivoExiste(caminhoArquivo, mensagemErro, incluirCaminhoNoErro) Then
        ImportarMolduraEObterGrupo = False
        Exit Function
    End If

    ActiveLayer.Import caminhoArquivo
    ImportarMolduraEObterGrupo = ObterGrupoImportado(grupoImportado)

End Function

Public Function ValidarCantoneiras( _
    ByVal cantSupDir As Shape, ByVal cantSupEsq As Shape, _
    ByVal cantInfEsq As Shape, ByVal cantInfDir As Shape) As Boolean

    If cantSupDir Is Nothing Or cantSupEsq Is Nothing _
    Or cantInfEsq Is Nothing Or cantInfDir Is Nothing Then
        MsgBox "Alguma cantoneira não foi encontrada no arquivo importado.", vbCritical
        ValidarCantoneiras = False
    Else
        ValidarCantoneiras = True
    End If

End Function

Public Function ValidarTubos( _
    ByVal tuboDir As Shape, ByVal tuboSup As Shape, _
    ByVal tuboEsq As Shape, ByVal tuboInf As Shape) As Boolean

    If tuboDir Is Nothing Or tuboSup Is Nothing _
    Or tuboEsq Is Nothing Or tuboInf Is Nothing Then
        MsgBox "Algum tubo não foi encontrado no arquivo importado.", vbCritical
        ValidarTubos = False
    Else
        ValidarTubos = True
    End If

End Function

Public Function ValidarAlhetasInferiores(ByVal alhetaInfEsq As Shape, ByVal alhetaInfDir As Shape) As Boolean
    ValidarAlhetasInferiores = Not (alhetaInfEsq Is Nothing Or alhetaInfDir Is Nothing)
End Function

Public Function ValidarArquivoExiste( _
    ByVal caminhoArquivo As String, _
    ByVal mensagem As String, _
    ByVal incluirCaminho As Boolean) As Boolean

    If Dir(caminhoArquivo) <> "" Then
        ValidarArquivoExiste = True
        Exit Function
    End If

    If incluirCaminho Then
        MsgBox mensagem & vbCrLf & caminhoArquivo, vbCritical
    Else
        MsgBox mensagem, vbCritical
    End If

    ValidarArquivoExiste = False

End Function

Public Function ObterGrupoImportado(ByRef grupo As Shape) As Boolean

    If ActiveSelectionRange.Count = 0 Then
        MsgBox "O arquivo foi importado, mas nenhum objeto ficou selecionado após a importação.", vbCritical
        ObterGrupoImportado = False
        Exit Function
    End If

    Set grupo = ActiveSelectionRange(1)

    If grupo Is Nothing Then
        MsgBox "Não foi possível obter o grupo importado.", vbCritical
        ObterGrupoImportado = False
        Exit Function
    End If

    If grupo.Shapes.Count = 0 Then
        MsgBox "O objeto importado não contém shapes internos.", vbCritical
        ObterGrupoImportado = False
        Exit Function
    End If

    ObterGrupoImportado = True

End Function

Public Function TentarObterRetanguloBaseMagenta(ByRef retanguloBase As Shape) As Boolean

    Dim shapePagina As Shape
    Dim retanguloSelecionado As Shape
    Dim quantidadeCandidatos As Long
    Dim maiorShape As Shape
    Dim maiorArea As Double
    Dim areaAtual As Double

    maiorArea = 0
    quantidadeCandidatos = 0

    For Each shapePagina In ActivePage.Shapes
        If shapePagina.Type = cdrRectangleShape Then
            If ShapeTemContornoMagenta(shapePagina) Then
                areaAtual = shapePagina.SizeWidth * shapePagina.SizeHeight
                If areaAtual > 1 Then
                    quantidadeCandidatos = quantidadeCandidatos + 1
                    If areaAtual > maiorArea Then
                        maiorArea = areaAtual
                        Set maiorShape = shapePagina
                    End If
                End If
            End If
        End If
    Next shapePagina

    If quantidadeCandidatos = 0 Then
        MsgBox "Nenhum retângulo com borda magenta encontrado.", vbExclamation
        TentarObterRetanguloBaseMagenta = False
        Exit Function
    End If

    If quantidadeCandidatos > 1 Then
        If ActiveSelectionRange.Count = 0 Then
            MsgBox "Mais de um retângulo magenta encontrado. Selecione manualmente o retângulo desejado.", vbCritical
            TentarObterRetanguloBaseMagenta = False
            Exit Function
        End If

        Set retanguloSelecionado = ActiveSelectionRange(1)

        If retanguloSelecionado.Type <> cdrRectangleShape Then
            MsgBox "O objeto selecionado não é um retângulo.", vbExclamation
            TentarObterRetanguloBaseMagenta = False
            Exit Function
        End If

        If Not ShapeTemContornoMagenta(retanguloSelecionado) Then
            MsgBox "O objeto selecionado não possui borda magenta CMYK válida.", vbExclamation
            TentarObterRetanguloBaseMagenta = False
            Exit Function
        End If

        Set retanguloBase = retanguloSelecionado
    Else
        Set retanguloBase = maiorShape
    End If

    TentarObterRetanguloBaseMagenta = True

End Function

Public Function ShapeTemContornoMagenta(ByVal s As Shape) As Boolean

    On Error GoTo Falha

    ShapeTemContornoMagenta = False
    If s Is Nothing Then Exit Function
    If s.Outline Is Nothing Then Exit Function

    ShapeTemContornoMagenta = CorEhMagentaCMYK(s.Outline.Color)
    Exit Function

Falha:
    ShapeTemContornoMagenta = False
    Err.Clear

End Function

Public Function CorEhMagentaCMYK(ByVal cor As Color) As Boolean

    On Error GoTo Falha

    CorEhMagentaCMYK = _
        Abs(cor.CMYKCyan - 0) < 0.5 And _
        Abs(cor.CMYKMagenta - 100) < 0.5 And _
        Abs(cor.CMYKYellow - 0) < 0.5 And _
        Abs(cor.CMYKBlack - 0) < 0.5

    Exit Function

Falha:
    CorEhMagentaCMYK = False
    Err.Clear

End Function


