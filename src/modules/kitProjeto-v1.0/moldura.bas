Attribute VB_Name = "molduraAutoV23"
Option Explicit

' ==============================================================================
' MÓDULO: Aplicação de Molduras - CorelDRAW VBA
' Descrição: Importa e posiciona molduras (azul, cinza, preto e economy) em
'            torno de um retângulo base com contorno magenta CMYK.
' ==============================================================================

' --- Caminhos dos arquivos de moldura ---
Private Const PASTA_MOLDURA_AUTO As String = _
    "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\MOLDURA AUTO\"

Private Const ARQUIVO_MOLDURA_AZUL    As String = PASTA_MOLDURA_AUTO & "molduraAzul.cdr"
Private Const ARQUIVO_MOLDURA_CINZA   As String = PASTA_MOLDURA_AUTO & "molduraCinza.cdr"
Private Const ARQUIVO_MOLDURA_PRETO   As String = PASTA_MOLDURA_AUTO & "molduraPreto.cdr"
Private Const ARQUIVO_MOLDURA_ECONOMY As String = PASTA_MOLDURA_AUTO & "molduraEconomy.cdr"

' --- Constantes de posicionamento (em milímetros) ---
Private Const OFFSET_MOLDURA_PADRAO_MM          As Double = 5.46
Private Const DESLOCAMENTO_MOLDURA_ECONOMY_MM   As Double = 6
Private Const DESLOCAMENTO_ALHETA_ECONOMY_MM    As Double = 55
Private Const LARGURA_MINIMA_DUPLICAR_ALHETA_MM As Double = 1830
Private Const LARGURA_MINIMA_DUPLICAR_ALHETA_EXTRA_MM As Double = 2000


' ==============================================================================
' MACROS PÚBLICAS — Chamadas pelo usuário ou painel de macros
' ==============================================================================

Sub molduraAzul()
    AplicarMoldura ARQUIVO_MOLDURA_AZUL
End Sub

Sub molduraCinza()
    AplicarMoldura ARQUIVO_MOLDURA_CINZA
End Sub

Sub molduraPreto()
    AplicarMoldura ARQUIVO_MOLDURA_PRETO
End Sub

Sub molduraEconomy()
    AplicarMolduraEconomy ARQUIVO_MOLDURA_ECONOMY
End Sub


' ==============================================================================
' APLICAÇÃO DE MOLDURA PADRÃO (Azul / Cinza / Preto)
' ==============================================================================

Private Sub AplicarMoldura(ByVal caminho As String)

    On Error GoTo TrataErro

    Dim retanguloBase As Shape
    Dim grupo As Shape
    Dim cantSupDir As Shape, cantSupEsq As Shape
    Dim cantInfEsq As Shape, cantInfDir As Shape
    Dim tuboDir As Shape, tuboSup As Shape
    Dim tuboEsq As Shape, tuboInf As Shape
    Dim offset As Double

    ' Converte o offset padrão para as unidades do documento
    offset = ActiveDocument.ToUnits(OFFSET_MOLDURA_PADRAO_MM, cdrMillimeter)

    If Not TentarObterRetanguloBaseMagenta(retanguloBase) Then Exit Sub
    If Not ValidarArquivoExiste(caminho, "Arquivo não encontrado: ", True) Then Exit Sub

    ' Importa o arquivo de moldura e obtém o grupo resultante
    ActiveLayer.Import caminho
    If Not ObterGrupoImportado(grupo) Then Exit Sub

    ' Mapeia as peças do grupo pelo nome
    MapearPecasBasicas grupo, cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
                              tuboDir, tuboSup, tuboEsq, tuboInf

    If Not ValidarCantoneiras(cantSupDir, cantSupEsq, cantInfEsq, cantInfDir) Then Exit Sub
    If Not ValidarTubos(tuboDir, tuboSup, tuboEsq, tuboInf) Then Exit Sub

    ' --- Posicionamento das cantoneiras nos cantos do retângulo base ---
    PosicionarCantoneiras retanguloBase, offset, _
                          cantSupDir, cantSupEsq, cantInfEsq, cantInfDir

    ' --- Posicionamento inicial dos tubos (antes do ajuste de tamanho) ---
    tuboDir.LeftX = cantSupDir.RightX - tuboDir.SizeWidth
    tuboDir.CenterY = retanguloBase.CenterY

    tuboSup.CenterX = retanguloBase.CenterX
    tuboSup.TopY = cantSupDir.TopY

    tuboEsq.LeftX = cantSupEsq.LeftX
    tuboEsq.CenterY = retanguloBase.CenterY

    tuboInf.CenterX = retanguloBase.CenterX
    tuboInf.BottomY = cantInfDir.BottomY

    ' --- Ajuste de largura/altura dos tubos conectando centros das cantoneiras ---
    AjustarTubosEntreCentos cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
                             tuboDir, tuboSup, tuboEsq, tuboInf

    Exit Sub

TrataErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "AplicarMoldura"

End Sub


' ==============================================================================
' APLICAÇÃO DE MOLDURA ECONOMY (com alhetas)
' ==============================================================================

Private Sub AplicarMolduraEconomy(ByVal caminhoArquivoMoldura As String)

    On Error GoTo TrataErro

    Dim retanguloBase As Shape
    Dim grupoMoldura As Shape

    ' Cantoneiras
    Dim cantSupDir As Shape, cantSupEsq As Shape
    Dim cantInfEsq As Shape, cantInfDir As Shape

    ' Tubos
    Dim tuboDir As Shape, tuboSup As Shape
    Dim tuboEsq As Shape, tuboInf As Shape

    ' Alhetas originais
    Dim alhetaInfDir As Shape, alhetaInfEsq As Shape
    Dim alhetaSupDir As Shape, alhetaSupEsq As Shape

    ' Alhetas duplicadas (para peças largas)
    Dim alhetaInfEsqDup     As Shape, alhetaSupEsqDup     As Shape
    Dim alhetaInfEsqDupExtra As Shape, alhetaInfDirDupExtra As Shape
    Dim alhetaSupEsqDupExtra As Shape, alhetaSupDirDupExtra As Shape

    ' Shapes de referência de alinhamento interno (dentro das cantoneiras)
    Dim refTuboDir As Shape, refTuboSup As Shape
    Dim refTuboEsq As Shape, refTuboInf As Shape

    Dim deslocamentoMoldura         As Double
    Dim deslocamentoHorizontalAlheta As Double

    deslocamentoMoldura = ActiveDocument.ToUnits(DESLOCAMENTO_MOLDURA_ECONOMY_MM, cdrMillimeter)
    deslocamentoHorizontalAlheta = ActiveDocument.ToUnits(DESLOCAMENTO_ALHETA_ECONOMY_MM, cdrMillimeter)

    If Not TentarObterRetanguloBaseMagenta(retanguloBase) Then Exit Sub
    If Not ValidarArquivoExiste(caminhoArquivoMoldura, "Arquivo não encontrado!", False) Then Exit Sub

    ' Importa o arquivo de moldura e obtém o grupo resultante
    ActiveLayer.Import caminhoArquivoMoldura
    If Not ObterGrupoImportado(grupoMoldura) Then Exit Sub

    ' Mapeia todas as peças (cantoneiras, tubos e alhetas)
    MapearPecasEconomy grupoMoldura, _
        cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
        tuboDir, tuboSup, tuboEsq, tuboInf, _
        alhetaInfDir, alhetaInfEsq, alhetaSupDir, alhetaSupEsq

    If Not ValidarCantoneiras(cantSupDir, cantSupEsq, cantInfEsq, cantInfDir) Then Exit Sub
    If Not ValidarTubos(tuboDir, tuboSup, tuboEsq, tuboInf) Then Exit Sub

    If alhetaInfEsq Is Nothing Or alhetaInfDir Is Nothing Then
        MsgBox "Alhetas inferiores não encontradas!", vbCritical
        Exit Sub
    End If

    ' Obtém shapes internos de referência para alinhamento dos tubos
    Set refTuboDir = BuscarShapePorNome(cantSupDir, "alinhaTuboDir")
    Set refTuboSup = BuscarShapePorNome(cantSupDir, "alinhaTuboSup")
    Set refTuboEsq = BuscarShapePorNome(cantInfEsq, "alinhaTuboEsq")
    Set refTuboInf = BuscarShapePorNome(cantInfEsq, "alinhaTuboInf")

    ' --- Posicionamento das cantoneiras com deslocamento da moldura economy ---
    PosicionarCantoneirasEconomy retanguloBase, deslocamentoMoldura, _
                                  cantSupDir, cantSupEsq, cantInfEsq, cantInfDir

    ' --- Posicionamento dos tubos (com fallback caso referência interna não exista) ---
    If Not refTuboDir Is Nothing Then
        tuboDir.CenterX = refTuboDir.CenterX
        tuboDir.CenterY = retanguloBase.CenterY
    Else
        tuboDir.LeftX = cantSupDir.RightX - tuboDir.SizeWidth
        tuboDir.CenterY = retanguloBase.CenterY
    End If

    tuboSup.CenterX = retanguloBase.CenterX
    If Not refTuboSup Is Nothing Then
        tuboSup.CenterY = refTuboSup.CenterY
    Else
        tuboSup.TopY = cantSupDir.TopY
    End If

    If Not refTuboEsq Is Nothing Then
        tuboEsq.CenterX = refTuboEsq.CenterX
        tuboEsq.CenterY = retanguloBase.CenterY
    Else
        tuboEsq.LeftX = cantSupEsq.LeftX
        tuboEsq.CenterY = retanguloBase.CenterY
    End If

    tuboInf.CenterX = retanguloBase.CenterX
    If Not refTuboInf Is Nothing Then
        tuboInf.CenterY = refTuboInf.CenterY
    Else
        tuboInf.BottomY = cantInfDir.BottomY
    End If

    ' --- Ajuste de largura/altura dos tubos conectando centros das cantoneiras ---
    AjustarTubosEntreCentos cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
                             tuboDir, tuboSup, tuboEsq, tuboInf

    ' --- Posicionamento das alhetas com deslocamento horizontal fixo ---
    alhetaInfEsq.LeftX = tuboInf.LeftX + deslocamentoHorizontalAlheta
    alhetaInfEsq.TopY = tuboInf.BottomY

    alhetaInfDir.LeftX = tuboInf.RightX - alhetaInfDir.SizeWidth - deslocamentoHorizontalAlheta
    alhetaInfDir.TopY = tuboInf.BottomY

    alhetaSupEsq.LeftX = tuboSup.LeftX + deslocamentoHorizontalAlheta
    alhetaSupEsq.BottomY = tuboSup.TopY

    alhetaSupDir.LeftX = tuboSup.RightX - alhetaSupDir.SizeWidth - deslocamentoHorizontalAlheta
    alhetaSupDir.BottomY = tuboSup.TopY

    ' --- Duplicação de alhetas para peças largas ---
    Dim larguraPeca As Double
    larguraPeca = retanguloBase.SizeWidth

    ' Entre 1830mm e 1999mm: adiciona apoio central somente nas alhetas esquerdas
    If larguraPeca >= ActiveDocument.ToUnits(LARGURA_MINIMA_DUPLICAR_ALHETA_MM, cdrMillimeter) _
    And larguraPeca < ActiveDocument.ToUnits(LARGURA_MINIMA_DUPLICAR_ALHETA_EXTRA_MM, cdrMillimeter) Then

        Set alhetaInfEsqDup = alhetaInfEsq.Duplicate
        Set alhetaSupEsqDup = alhetaSupEsq.Duplicate

        ' Centraliza as duplicatas horizontalmente no eixo dos tubos
        alhetaInfEsqDup.CenterX = tuboInf.CenterX
        alhetaSupEsqDup.CenterX = tuboSup.CenterX

    End If

    ' A partir de 2000mm: distribui 2 duplicatas entre o centro e as cantoneiras
    If larguraPeca >= ActiveDocument.ToUnits(LARGURA_MINIMA_DUPLICAR_ALHETA_EXTRA_MM, cdrMillimeter) Then

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


' ==============================================================================
' POSICIONAMENTO DE CANTONEIRAS
' ==============================================================================

' Posiciona as 4 cantoneiras com offset simétrico para molduras padrão.
Private Sub PosicionarCantoneiras(ByVal base As Shape, ByVal offset As Double, _
    ByVal cantSupDir As Shape, ByVal cantSupEsq As Shape, _
    ByVal cantInfEsq As Shape, ByVal cantInfDir As Shape)

    cantSupDir.LeftX = base.RightX - cantSupDir.SizeWidth + offset
    cantSupDir.TopY = base.TopY + offset

    cantSupEsq.LeftX = base.LeftX - offset
    cantSupEsq.TopY = base.TopY + offset

    cantInfEsq.LeftX = base.LeftX - offset
    cantInfEsq.TopY = base.BottomY + cantInfEsq.SizeHeight - offset

    cantInfDir.LeftX = base.RightX - cantInfDir.SizeWidth + offset
    cantInfDir.TopY = base.BottomY + cantInfDir.SizeHeight - offset

End Sub

' Posiciona as 4 cantoneiras com deslocamento específico da moldura economy.
Private Sub PosicionarCantoneirasEconomy(ByVal base As Shape, ByVal desl As Double, _
    ByVal cantSupDir As Shape, ByVal cantSupEsq As Shape, _
    ByVal cantInfEsq As Shape, ByVal cantInfDir As Shape)

    cantSupDir.LeftX = base.RightX - cantSupDir.SizeWidth + desl
    cantSupDir.TopY = base.TopY + desl

    cantSupEsq.LeftX = base.LeftX - desl
    cantSupEsq.TopY = base.TopY + desl

    cantInfEsq.LeftX = base.LeftX - desl
    cantInfEsq.TopY = base.BottomY + cantInfEsq.SizeHeight - desl

    cantInfDir.LeftX = base.RightX - cantInfDir.SizeWidth + desl
    cantInfDir.TopY = base.BottomY + cantInfDir.SizeHeight - desl

End Sub


' ==============================================================================
' AJUSTE DE TAMANHO DOS TUBOS
' Redimensiona cada tubo para conectar os centros das cantoneiras opostas.
' ==============================================================================

Private Sub AjustarTubosEntreCentos( _
    ByVal cantSupDir As Shape, ByVal cantSupEsq As Shape, _
    ByVal cantInfEsq As Shape, ByVal cantInfDir As Shape, _
    ByVal tuboDir As Shape, ByVal tuboSup As Shape, _
    ByVal tuboEsq As Shape, ByVal tuboInf As Shape)

    Dim largura As Double
    Dim altura   As Double

    ' Tubo inferior: entre os centros X das cantoneiras inferiores
    largura = cantInfDir.CenterX - cantInfEsq.CenterX
    tuboInf.SetSize largura, tuboInf.SizeHeight
    tuboInf.CenterX = (cantInfEsq.CenterX + cantInfDir.CenterX) / 2

    ' Tubo esquerdo: entre os centros Y das cantoneiras esquerdas
    altura = cantSupEsq.CenterY - cantInfEsq.CenterY
    tuboEsq.SetSize tuboEsq.SizeWidth, altura
    tuboEsq.CenterY = (cantSupEsq.CenterY + cantInfEsq.CenterY) / 2

    ' Tubo direito: entre os centros Y das cantoneiras direitas
    altura = cantSupDir.CenterY - cantInfDir.CenterY
    tuboDir.SetSize tuboDir.SizeWidth, altura
    tuboDir.CenterY = (cantSupDir.CenterY + cantInfDir.CenterY) / 2

    ' Tubo superior: entre os centros X das cantoneiras superiores
    largura = cantSupDir.CenterX - cantSupEsq.CenterX
    tuboSup.SetSize largura, tuboSup.SizeHeight
    tuboSup.CenterX = (cantSupEsq.CenterX + cantSupDir.CenterX) / 2

End Sub


' ==============================================================================
' VALIDAÇÕES
' ==============================================================================

' Retorna True se todas as cantoneiras foram encontradas no grupo.
Private Function ValidarCantoneiras( _
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

' Retorna True se todos os tubos foram encontrados no grupo.
Private Function ValidarTubos( _
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

' Verifica se o arquivo existe no caminho informado.
Private Function ValidarArquivoExiste( _
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

' Tenta obter o objeto importado após ActiveLayer.Import.
' Retorna False e exibe mensagem se nenhum objeto foi selecionado.
Private Function ObterGrupoImportado(ByRef grupo As Shape) As Boolean

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

' Localiza o maior retângulo com contorno magenta CMYK na página ativa.
' Se houver mais de um candidato, exige que o usuário selecione manualmente.
Private Function TentarObterRetanguloBaseMagenta(ByRef retanguloBase As Shape) As Boolean

    Dim shapePagina       As Shape
    Dim retanguloSelecionado As Shape
    Dim candidatos        As Collection
    Dim maiorShape        As Shape
    Dim maiorArea         As Double
    Dim areaAtual         As Double

    Set candidatos = New Collection
    maiorArea = 0

    For Each shapePagina In ActivePage.Shapes
        If shapePagina.Type = cdrRectangleShape Then
            If ShapeTemContornoMagenta(shapePagina) Then
                areaAtual = shapePagina.SizeWidth * shapePagina.SizeHeight
                If areaAtual > 1 Then
                    candidatos.Add shapePagina
                    If areaAtual > maiorArea Then
                        maiorArea = areaAtual
                        Set maiorShape = shapePagina
                    End If
                End If
            End If
        End If
    Next shapePagina

    If candidatos.Count = 0 Then
        MsgBox "Nenhum retângulo com borda magenta encontrado.", vbExclamation
        TentarObterRetanguloBaseMagenta = False
        Exit Function
    End If

    ' Com mais de um candidato, exige seleção manual do usuário
    If candidatos.Count > 1 Then
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


' ==============================================================================
' MAPEAMENTO DE PEÇAS POR NOME
' ==============================================================================

' Percorre os shapes do grupo e associa cada peça pelo nome esperado.
Private Sub MapearPecasBasicas(ByVal grupo As Shape, _
    ByRef cantSupDir As Shape, ByRef cantSupEsq As Shape, _
    ByRef cantInfEsq As Shape, ByRef cantInfDir As Shape, _
    ByRef tuboDir As Shape, ByRef tuboSup As Shape, _
    ByRef tuboEsq As Shape, ByRef tuboInf As Shape)

    Dim s As Shape

    For Each s In grupo.Shapes
        Select Case s.Name
            Case "cantSupDir": Set cantSupDir = s
            Case "cantSupEsq": Set cantSupEsq = s
            Case "cantInfEsq": Set cantInfEsq = s
            Case "cantInfDir": Set cantInfDir = s
            Case "tuboDir":    Set tuboDir = s
            Case "tuboSup":    Set tuboSup = s
            Case "tuboEsq":    Set tuboEsq = s
            Case "tuboInf":    Set tuboInf = s
        End Select
    Next s

End Sub

' Percorre os shapes do grupo e associa cantoneiras, tubos e alhetas pelo nome.
Private Sub MapearPecasEconomy(ByVal grupo As Shape, _
    ByRef cantSupDir As Shape, ByRef cantSupEsq As Shape, _
    ByRef cantInfEsq As Shape, ByRef cantInfDir As Shape, _
    ByRef tuboDir As Shape, ByRef tuboSup As Shape, _
    ByRef tuboEsq As Shape, ByRef tuboInf As Shape, _
    ByRef alhetaInfDir As Shape, ByRef alhetaInfEsq As Shape, _
    ByRef alhetaSupDir As Shape, ByRef alhetaSupEsq As Shape)

    Dim s As Shape

    For Each s In grupo.Shapes
        Select Case s.Name
            Case "cantSupDir":   Set cantSupDir = s
            Case "cantSupEsq":   Set cantSupEsq = s
            Case "cantInfEsq":   Set cantInfEsq = s
            Case "cantInfDir":   Set cantInfDir = s
            Case "tuboDir":      Set tuboDir = s
            Case "tuboSup":      Set tuboSup = s
            Case "tuboEsq":      Set tuboEsq = s
            Case "tuboInf":      Set tuboInf = s
            Case "alhetaInfDir": Set alhetaInfDir = s
            Case "alhetaInfEsq": Set alhetaInfEsq = s
            Case "alhetaSupDir": Set alhetaSupDir = s
            Case "alhetaSupEsq": Set alhetaSupEsq = s
        End Select
    Next s

End Sub


' ==============================================================================
' BUSCA RECURSIVA POR NOME
' ==============================================================================

' Percorre grupos aninhados em busca de um shape pelo nome.
' Retorna Nothing se não encontrado.
Private Function BuscarShapePorNome(ByVal grp As Shape, ByVal nome As String) As Shape

    Dim s         As Shape
    Dim resultado As Shape

    If grp.Type <> cdrGroupShape Then Exit Function

    For Each s In grp.Shapes

        If s.Name = nome Then
            Set BuscarShapePorNome = s
            Exit Function
        End If

        If s.Type = cdrGroupShape Then
            Set resultado = BuscarShapePorNome(s, nome)
            If Not resultado Is Nothing Then
                Set BuscarShapePorNome = resultado
                Exit Function
            End If
        End If

    Next s

End Function


' ==============================================================================
' DETECÇÃO DE COR MAGENTA CMYK
' ==============================================================================

' Retorna True se o shape possui contorno com cor magenta CMYK pura.
Private Function ShapeTemContornoMagenta(ByVal s As Shape) As Boolean

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

' Retorna True se a cor CMYK corresponde ao magenta puro (0, 100, 0, 0).
Private Function CorEhMagentaCMYK(ByVal cor As Color) As Boolean

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

