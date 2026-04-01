Attribute VB_Name = "cavaleteAutoV2"
Option Explicit

' =========================================================
' CONFIGURAÇÕES GERAIS
' =========================================================

' Margem de tolerância usada para validar a cor magenta em CMYK.
Private Const TOLERANCIA_COR As Double = 0.5

' Caminhos dos arquivos de cavalete que serão importados.
Private Const CAMINHO_CAVALETE_CZ As String = "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\CAVALETES\CAVALETE_CZ.cdr"
Private Const CAMINHO_CAVALETE_BR As String = "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\CAVALETES\CAVALETE_BR.cdr"
Private Const CAMINHO_CAVALETE_PT As String = "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\CAVALETES\CAVALETE_PT.cdr"

' Nome do grupo esperado dentro de cada arquivo importado.
Private Const NOME_GRUPO_CZ As String = "CAVALETE-METALON3-CZ"
Private Const NOME_GRUPO_BR As String = "CAVALETE-METALON3-BR"
Private Const NOME_GRUPO_PT As String = "CAVALETE-METALON3-PT"

' Deslocamentos usados no posicionamento dos objetos no documento.
Private Const DESLOCAMENTO_X_CAVALETE_MM As Double = 418.8
Private Const DESLOCAMENTO_Y_CAVALETE_MM As Double = 30.4
Private Const DESLOCAMENTO_Y_MAO_FRANCESA_MM As Double = 188.419
Private Const DESLOCAMENTO_X_GRUPO_ESPELHADO_MM As Double = 147

' ROTINAS PÚBLICAS

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

' ROTINA PRINCIPAL

' Controla todo o fluxo:
' 1) obtém o quadro magenta
' 2) importa o cavalete
' 3) posiciona o cavalete
' 4) encontra o grupo correto
' 5) encontra a mão francesa
' 6) posiciona a mão francesa
' 7) duplica, espelha e posiciona o grupo

Private Sub InserirCavalete(ByVal caminhoArquivo As String, ByVal nomeGrupo As String)
    Dim quadro As Shape
    Dim shapeImportado As Shape
    Dim grupoCavalete As Shape
    Dim maoFrancesa As Shape
    Dim grupoEspelhado As Shape

    On Error GoTo TrataErro

    Set quadro = ObterQuadroMagentaValido()
    If quadro Is Nothing Then Exit Sub

    If Not ArquivoExiste(caminhoArquivo) Then
        MsgBox "Arquivo não encontrado:" & vbCrLf & caminhoArquivo, vbCritical
        Exit Sub
    End If

    Set shapeImportado = ImportarArquivoCavalete(caminhoArquivo)
    If shapeImportado Is Nothing Then Exit Sub

    PosicionarCavaleteInicial shapeImportado, quadro

    Set grupoCavalete = ObterGrupoPorNome(shapeImportado, nomeGrupo)
    If grupoCavalete Is Nothing Then
        MsgBox "Grupo '" & nomeGrupo & "' não encontrado.", vbCritical
        Exit Sub
    End If

    Set maoFrancesa = BuscarShapePorNomeRecursivo(grupoCavalete, "maoFrancesa")
    If maoFrancesa Is Nothing Then
        MsgBox "Objeto 'maoFrancesa' não encontrado dentro do grupo '" & nomeGrupo & "'.", vbCritical
        Exit Sub
    End If

    PosicionarMaoFrancesa maoFrancesa, quadro

    Set grupoEspelhado = grupoCavalete.Duplicate
    EspelharEPosicionarGrupo grupoEspelhado, quadro

    Exit Sub

TrataErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' LOCALIZAÇÃO DO QUADRO MAGENTA

' Retorna o quadro magenta válido que será usado como referência.
' Se existir mais de um, tenta usar o objeto selecionado pelo usuário.
Private Function ObterQuadroMagentaValido() As Shape
    Dim candidatos As Collection
    Dim maiorRetangulo As Shape
    Dim shapeSelecionado As Shape

    Set candidatos = New Collection
    Set maiorRetangulo = BuscarMaiorRetanguloMagenta(candidatos)

    If candidatos.Count = 0 Then
        MsgBox "Nenhum retângulo com borda magenta encontrado.", vbExclamation
        Exit Function
    End If

    If candidatos.Count = 1 Then
        Set ObterQuadroMagentaValido = maiorRetangulo
        Exit Function
    End If

    If ActiveSelection.Shapes.Count > 0 Then
        Set shapeSelecionado = ActiveSelection.Shapes(1)

        If Not EhRetanguloMagenta(shapeSelecionado) Then
            MsgBox "O objeto selecionado não é um retângulo com borda magenta.", vbExclamation
            Exit Function
        End If

        Set ObterQuadroMagentaValido = shapeSelecionado
    Else
        MsgBox "Mais de um retângulo com borda magenta encontrado. Selecione manualmente o quadro.", vbCritical
    End If
End Function

' Procura todos os retângulos com borda magenta da página
' e retorna o maior deles por área.
Private Function BuscarMaiorRetanguloMagenta(ByRef candidatos As Collection) As Shape
    Dim s As Shape
    Dim maiorShape As Shape
    Dim maiorArea As Double
    Dim areaAtual As Double

    maiorArea = 0

    For Each s In ActivePage.Shapes
        If EhRetanguloMagenta(s) Then
            areaAtual = s.SizeWidth * s.SizeHeight

            If areaAtual > 1 Then
                candidatos.Add s

                If areaAtual > maiorArea Then
                    maiorArea = areaAtual
                    Set maiorShape = s
                End If
            End If
        End If
    Next s

    Set BuscarMaiorRetanguloMagenta = maiorShape
End Function

' Valida se uma shape é um retângulo com borda magenta.
Private Function EhRetanguloMagenta(ByVal s As Shape) As Boolean
    On Error GoTo Falha

    EhRetanguloMagenta = False

    If s Is Nothing Then Exit Function
    If s.Type <> cdrRectangleShape Then Exit Function
    If s.Outline Is Nothing Then Exit Function

    If Abs(s.Outline.Color.CMYKCyan - 0) < TOLERANCIA_COR And _
       Abs(s.Outline.Color.CMYKMagenta - 100) < TOLERANCIA_COR And _
       Abs(s.Outline.Color.CMYKYellow - 0) < TOLERANCIA_COR And _
       Abs(s.Outline.Color.CMYKBlack - 0) < TOLERANCIA_COR Then
        EhRetanguloMagenta = True
    End If

    Exit Function

Falha:
    EhRetanguloMagenta = False
End Function

' IMPORTAÇÃO

' Importa o arquivo do cavalete e retorna a shape selecionada após a importação.
Private Function ImportarArquivoCavalete(ByVal caminhoArquivo As String) As Shape
    ActiveLayer.Import caminhoArquivo

    If ActiveSelection Is Nothing Then
        MsgBox "Falha ao importar o cavalete: nenhuma seleção ativa foi criada.", vbCritical
        Exit Function
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Falha ao importar o cavalete: nenhum objeto foi selecionado após importar.", vbCritical
        Exit Function
    End If

    Set ImportarArquivoCavalete = ActiveSelection.Shapes(1)
End Function

' POSICIONAMENTO

' Posiciona o cavalete importado tomando o quadro como referência.
Private Sub PosicionarCavaleteInicial(ByVal cavalete As Shape, ByVal quadro As Shape)
    cavalete.TopY = quadro.TopY
    cavalete.LeftX = quadro.LeftX

    cavalete.LeftX = cavalete.LeftX - MmParaDocumento(DESLOCAMENTO_X_CAVALETE_MM)
    cavalete.TopY = cavalete.TopY + MmParaDocumento(DESLOCAMENTO_Y_CAVALETE_MM)
End Sub

' Posiciona a mão francesa usando a base inferior do quadro como referência.
Private Sub PosicionarMaoFrancesa(ByVal maoFrancesa As Shape, ByVal quadro As Shape)
    maoFrancesa.BottomY = quadro.BottomY
    maoFrancesa.BottomY = maoFrancesa.BottomY - MmParaDocumento(DESLOCAMENTO_Y_MAO_FRANCESA_MM)
End Sub

' Espelha horizontalmente o grupo duplicado e o posiciona do lado direito do quadro.
Private Sub EspelharEPosicionarGrupo(ByVal grupo As Shape, ByVal quadro As Shape)
    grupo.Flip cdrFlipHorizontal
    grupo.RightX = quadro.RightX
    grupo.RightX = grupo.RightX + MmParaDocumento(DESLOCAMENTO_X_GRUPO_ESPELHADO_MM)
End Sub

' BUSCA DE SHAPES E GRUPOS

' Procura um grupo pelo nome dentro da shape raiz importada.
Private Function ObterGrupoPorNome(ByVal shapeRaiz As Shape, ByVal nomeGrupo As String) As Shape
    Dim s As Shape

    On Error GoTo Falha

    For Each s In shapeRaiz.Shapes.All
        If s.Type = cdrGroupShape Then
            If NomesIguais(s.Name, nomeGrupo) Then
                Set ObterGrupoPorNome = s
                Exit Function
            End If
        End If
    Next s

    Exit Function

Falha:
    Set ObterGrupoPorNome = Nothing
End Function

' Busca uma shape recursivamente pelo nome, começando na raiz informada.
Private Function BuscarShapePorNomeRecursivo(ByVal raiz As Shape, ByVal nomeBuscado As String) As Shape
    Dim s As Shape
    Dim encontrado As Shape

    On Error GoTo Falha

    If raiz Is Nothing Then Exit Function

    If NomesIguais(raiz.Name, nomeBuscado) Then
        Set BuscarShapePorNomeRecursivo = raiz
        Exit Function
    End If

    If Not TemFilhos(raiz) Then Exit Function

    For Each s In raiz.Shapes.All
        If NomesIguais(s.Name, nomeBuscado) Then
            Set BuscarShapePorNomeRecursivo = s
            Exit Function
        End If

        If TemFilhos(s) Then
            Set encontrado = BuscarShapePorNomeRecursivo(s, nomeBuscado)
            If Not encontrado Is Nothing Then
                Set BuscarShapePorNomeRecursivo = encontrado
                Exit Function
            End If
        End If
    Next s

    Exit Function

Falha:
    Set BuscarShapePorNomeRecursivo = Nothing
End Function

' Informa se a shape possui objetos filhos.
Private Function TemFilhos(ByVal s As Shape) As Boolean
    On Error GoTo Falha

    TemFilhos = False

    If s Is Nothing Then Exit Function
    TemFilhos = (s.Shapes.Count > 0)

    Exit Function

Falha:
    TemFilhos = False
End Function

' Compara dois nomes ignorando espaços extras e diferença de maiúsculas/minúsculas.
Private Function NomesIguais(ByVal nome1 As String, ByVal nome2 As String) As Boolean
    NomesIguais = (StrComp(Trim$(nome1), Trim$(nome2), vbTextCompare) = 0)
End Function

' UTILITÁRIOS

' Converte um valor em milímetros para a unidade atual do documento.
Private Function MmParaDocumento(ByVal valorMm As Double) As Double
    MmParaDocumento = ActiveDocument.ToUnits(valorMm, cdrMillimeter)
End Function

' Verifica se o arquivo informado existe no caminho especificado.
Private Function ArquivoExiste(ByVal caminhoArquivo As String) As Boolean
    ArquivoExiste = (Dir$(caminhoArquivo) <> "")
End Function

