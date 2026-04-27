Attribute VB_Name = "modDescricao"
' modDescricao
Option Explicit

Private Const SHAPE_KSVR_A4_AD As String = "KSVR-A4-AD-MACRO"
Private Const SHAPE_KSVR_A4_MG As String = "KSVR-A4-MG-MACRO"
Private Const SHAPE_KSVP_A4_AD As String = "KSVP-A4-AD-MACRO"
Private Const SHAPE_KSVP_A4_MG As String = "KSVP-A4-MG-MACRO"
Private Const SHAPE_TESTEIRA As String = "TESTEIRA-MACRO"
Private Const SHAPE_DAVN As String = "DAVN-MACRO"
Private Const SHAPE_ESC_A4_CZ As String = "ESC-A4-CZ-MACRO"
Private Const SHAPE_ESC_A4_AM As String = "ESC-A4-AM-MACRO"
Private Const SHAPE_ESC_A4_AZ As String = "ESC-A4-AZ-MACRO"
Private Const SHAPE_ESC_A4_VD As String = "ESC-A4-VD-MACRO"
Private Const SHAPE_ESC_A4_VM As String = "ESC-A4-VM-MACRO"
Private Const SHAPE_ESC_A4_PT As String = "ESC-A4-PT-MACRO"
Private Const SHAPE_BASE_ESC_A4 As String = "BASE-ESC-A4-MACRO"

Public Function MontarTextoCompleto(ByVal tipo As tipoQuadro, _
                                     ByVal altura As Double, _
                                     ByVal largura As Double, _
                                     ByVal catalogo As Collection, _
                                     ByVal contadores As Object, _
                                     ByVal medidasAcessorios As Object) As String
    Dim texto As String
    texto = MontarTextoPrincipal(tipo, altura, largura)

    Dim secaoAcessorios As String
    secaoAcessorios = MontarTextoAcessorios(tipo, catalogo, contadores, medidasAcessorios)

    MontarTextoCompleto = AnexarSecaoAcessorios(texto, secaoAcessorios)
End Function

Private Function MontarTextoPrincipal(ByVal tipo As tipoQuadro, _
                                       ByVal altura As Double, _
                                       ByVal largura As Double) As String
    Select Case tipo
        Case tqQPMM_P
            MontarTextoPrincipal = "QUADRO BRANCO MAGNÉTICO" & vbCrLf & _
                                   "PARA ESCRITA COM IMPRESSÃO " & vbCrLf & _
                                   "DIGITAL UV. E LAMINAÇÃO PYT" & vbCrLf & _
                                   "MED " & altura & "x" & largura & "MM - QPMM"
        Case tqQBTA
            MontarTextoPrincipal = "QUADRO BRANCO MEDINDO" & vbCrLf & _
                                   "(" & altura & "X" & largura & ")MM - QBTA"
        Case Else
            MontarTextoPrincipal = "QUADRO BRANCO PARA ESCRITA" & vbCrLf & _
                                   "COM IMPRESSÃO DIGITAL UV. E" & vbCrLf & _
                                   "LAMINAÇÃO PYT MED " & altura & "x" & largura & "MM" & vbCrLf & _
                                   "- QPMS"
    End Select
End Function

Private Function MontarTextoAcessorios(ByVal tipo As tipoQuadro, _
                                        ByVal catalogo As Collection, _
                                        ByVal contadores As Object, _
                                        ByVal medidasAcessorios As Object) As String
    Dim item As Variant
    Dim nomeShape As String
    Dim quantidade As Long
    Dim texto As String
    Dim outputCode As String
    Dim linhaEscAgrupada As String
    Dim escLinhaInserida As Boolean

    texto = ""
    linhaEscAgrupada = MontarLinhaEscA4Agrupada(contadores)

    For Each item In catalogo
        nomeShape = CStr(item("ShapeName"))
        If linhaEscAgrupada <> "" And EhShapeEscA4(nomeShape) Then
            If Not escLinhaInserida Then
                texto = texto & linhaEscAgrupada
                escLinhaInserida = True
            End If
        Else
            quantidade = CLng(contadores(nomeShape))
            If quantidade > 0 Then
                outputCode = CStr(item("OutputCode"))
                If EhAcessorioComMedidaSeparada(nomeShape) Then
                    texto = texto & MontarLinhasAcessorioComMedida(tipo, nomeShape, quantidade, outputCode, medidasAcessorios)
                ElseIf EhAcessorioComVarianteBorda(nomeShape) Then
                    outputCode = ResolverOutputCode(tipo, nomeShape, outputCode, medidasAcessorios)
                    texto = texto & MontarLinhasAcessorioComVariante(nomeShape, quantidade, outputCode, medidasAcessorios)
                ElseIf EhCavaleteMetalon3(nomeShape) Then
                    texto = texto & MontarLinhaSemQuantidade(outputCode)
                Else
                    outputCode = ResolverOutputCode(tipo, nomeShape, outputCode, medidasAcessorios)
                    outputCode = AjustarPluralBaseEscA4(nomeShape, quantidade, outputCode)
                    texto = texto & MontarLinhaComQuantidade(quantidade, outputCode)
                End If
            End If
        End If
    Next item

    texto = texto & MontarLinhasKanbanPorGrupo(medidasAcessorios)

    MontarTextoAcessorios = texto
End Function

Private Function MontarLinhaEscA4Agrupada(ByVal contadores As Object) As String
    Dim total As Long
    Dim qtdCores As Long
    Dim detalhe As String
    Dim qtd As Long

    total = 0
    qtdCores = 0
    detalhe = ""

    qtd = ObterQtdShape(contadores, SHAPE_ESC_A4_CZ)
    If qtd > 0 Then
        total = total + qtd
        qtdCores = qtdCores + 1
        AdicionarDetalheEsc detalhe, qtd, "CZ", qtdCores
    End If

    qtd = ObterQtdShape(contadores, SHAPE_ESC_A4_AM)
    If qtd > 0 Then
        total = total + qtd
        qtdCores = qtdCores + 1
        AdicionarDetalheEsc detalhe, qtd, "AM", qtdCores
    End If

    qtd = ObterQtdShape(contadores, SHAPE_ESC_A4_AZ)
    If qtd > 0 Then
        total = total + qtd
        qtdCores = qtdCores + 1
        AdicionarDetalheEsc detalhe, qtd, "AZ", qtdCores
    End If

    qtd = ObterQtdShape(contadores, SHAPE_ESC_A4_VD)
    If qtd > 0 Then
        total = total + qtd
        qtdCores = qtdCores + 1
        AdicionarDetalheEsc detalhe, qtd, "VD", qtdCores
    End If

    qtd = ObterQtdShape(contadores, SHAPE_ESC_A4_VM)
    If qtd > 0 Then
        total = total + qtd
        qtdCores = qtdCores + 1
        AdicionarDetalheEsc detalhe, qtd, "VM", qtdCores
    End If

    qtd = ObterQtdShape(contadores, SHAPE_ESC_A4_PT)
    If qtd > 0 Then
        total = total + qtd
        qtdCores = qtdCores + 1
        AdicionarDetalheEsc detalhe, qtd, "PT", qtdCores
    End If

    If qtdCores >= 2 Then
        MontarLinhaEscA4Agrupada = "- " & total & " ESC A4 (" & detalhe & ")" & vbCrLf
    End If
End Function

Private Function ObterQtdShape(ByVal contadores As Object, _
                               ByVal nomeShape As String) As Long
    If contadores Is Nothing Then Exit Function
    If Not contadores.Exists(nomeShape) Then Exit Function

    ObterQtdShape = CLng(contadores(nomeShape))
End Function

Private Sub AdicionarDetalheEsc(ByRef detalhe As String, _
                                ByVal quantidade As Long, _
                                ByVal cor As String, _
                                ByVal indiceCor As Long)
    If detalhe <> "" Then detalhe = detalhe & ","
    If indiceCor = 4 Then detalhe = detalhe & vbCrLf
    detalhe = detalhe & quantidade & " " & cor
End Sub

Private Function EhShapeEscA4(ByVal nomeShape As String) As Boolean
    Select Case UCase$(nomeShape)
        Case SHAPE_ESC_A4_CZ, SHAPE_ESC_A4_AM, SHAPE_ESC_A4_AZ, _
             SHAPE_ESC_A4_VD, SHAPE_ESC_A4_VM, SHAPE_ESC_A4_PT
            EhShapeEscA4 = True
    End Select
End Function

Private Function AjustarPluralBaseEscA4(ByVal nomeShape As String, _
                                        ByVal quantidade As Long, _
                                        ByVal outputCode As String) As String
    Dim saida As String

    AjustarPluralBaseEscA4 = outputCode

    If UCase$(nomeShape) <> SHAPE_BASE_ESC_A4 Then Exit Function
    If quantidade <= 1 Then Exit Function

    saida = outputCode
    saida = Replace(saida, "BASE-ESC-A4", "BASES-ESC-A4")
    saida = Replace(saida, "BASE ESC A4", "BASES ESC A4")

    AjustarPluralBaseEscA4 = saida
End Function

Private Function AnexarSecaoAcessorios(ByVal textoPrincipal As String, _
                                        ByVal textoAcessorios As String) As String
    If textoAcessorios = "" Then
        AnexarSecaoAcessorios = textoPrincipal
    Else
        AnexarSecaoAcessorios = textoPrincipal & vbCrLf & vbCrLf & _
                                "ACESSÓRIOS:" & vbCrLf & vbCrLf & _
                                textoAcessorios
    End If
End Function

Private Function EhCavaleteMetalon3(ByVal nomeShape As String) As Boolean
    EhCavaleteMetalon3 = (Left$(UCase$(nomeShape), 18) = "CAVALETE-METALON3-")
End Function

Private Function ResolverOutputCode(ByVal tipo As tipoQuadro, _
                                    ByVal nomeShape As String, _
                                    ByVal outputCode As String, _
                                    ByVal medidasAcessorios As Object, _
                                    Optional ByVal medidaOverride As String = "") As String
    Dim nomeShapeNormalizado As String
    Dim sufixoCompat As String

    nomeShapeNormalizado = UCase$(nomeShape)

    If nomeShapeNormalizado = SHAPE_DAVN Then
        If tipo = tqQPMM_P Then
            sufixoCompat = COMPAT_MG
        Else
            sufixoCompat = COMPAT_AD
        End If
        outputCode = Replace(outputCode, "TIPO", sufixoCompat)
    End If

    If nomeShapeNormalizado = SHAPE_TESTEIRA Or nomeShapeNormalizado = SHAPE_DAVN Then
        If medidaOverride <> "" Then
            ResolverOutputCode = Replace(outputCode, "ALTXLARGURA", medidaOverride)
            Exit Function
        End If

        If Not medidasAcessorios Is Nothing Then
            If medidasAcessorios.Exists(nomeShape) Then
                ResolverOutputCode = Replace(outputCode, "ALTXLARGURA", CStr(medidasAcessorios(nomeShape)))
                Exit Function
            End If
        End If
    End If

    ResolverOutputCode = outputCode
End Function

Private Function EhAcessorioComVarianteBorda(ByVal nomeShape As String) As Boolean
    Select Case UCase$(nomeShape)
        Case SHAPE_KSVR_A4_AD, SHAPE_KSVR_A4_MG, SHAPE_KSVP_A4_AD, SHAPE_KSVP_A4_MG
            EhAcessorioComVarianteBorda = True
    End Select
End Function

Private Function MontarLinhasAcessorioComVariante(ByVal nomeShape As String, _
                                                  ByVal quantidadeTotal As Long, _
                                                  ByVal outputCodePadrao As String, _
                                                  ByVal medidasAcessorios As Object) As String
    Dim qtdUniforme As Long
    Dim qtdDegrade As Long
    Dim qtdComVariante As Long
    Dim qtdSemVariante As Long
    Dim texto As String
    Dim nomeBase As String

    texto = ""
    nomeBase = NomeBasePorShape(nomeShape)
    qtdUniforme = ObterQuantidadeVariante(medidasAcessorios, nomeShape, "UNIFORME")
    qtdDegrade = ObterQuantidadeVariante(medidasAcessorios, nomeShape, "DEGRADÊ")

    If qtdUniforme > 0 Then
        texto = texto & "- " & qtdUniforme & " " & nomeBase & " UNIFORME" & vbCrLf
    End If

    If qtdDegrade > 0 Then
        texto = texto & "- " & qtdDegrade & " " & nomeBase & " DEGRADÊ" & vbCrLf
    End If

    qtdComVariante = qtdUniforme + qtdDegrade
    qtdSemVariante = quantidadeTotal - qtdComVariante
    If qtdSemVariante > 0 Then
        texto = texto & "- " & qtdSemVariante & " " & outputCodePadrao & vbCrLf
    End If

    MontarLinhasAcessorioComVariante = texto
End Function

Private Function ObterQuantidadeVariante(ByVal medidasAcessorios As Object, _
                                         ByVal nomeShape As String, _
                                         ByVal variante As String) As Long
    Dim chave As String
    chave = UCase$(nomeShape) & "_VARIANTE_" & UCase$(variante) & "_QTD"

    If medidasAcessorios Is Nothing Then Exit Function
    If medidasAcessorios.Exists(chave) Then
        ObterQuantidadeVariante = CLng(medidasAcessorios(chave))
    End If
End Function

Private Function NomeBasePorShape(ByVal nomeShape As String) As String
    Select Case UCase$(nomeShape)
        Case SHAPE_KSVR_A4_AD
            NomeBasePorShape = "KSVR-A4-AD"
        Case SHAPE_KSVR_A4_MG
            NomeBasePorShape = "KSVR-A4-MG"
        Case SHAPE_KSVP_A4_AD
            NomeBasePorShape = "KSVP-A4-AD"
        Case SHAPE_KSVP_A4_MG
            NomeBasePorShape = "KSVP-A4-MG"
    End Select
End Function

Private Function EhAcessorioComMedidaSeparada(ByVal nomeShape As String) As Boolean
    Select Case UCase$(nomeShape)
        Case SHAPE_TESTEIRA, SHAPE_DAVN
            EhAcessorioComMedidaSeparada = True
    End Select
End Function

Private Function MontarLinhasAcessorioComMedida(ByVal tipo As tipoQuadro, _
                                                 ByVal nomeShape As String, _
                                                 ByVal quantidadeTotal As Long, _
                                                 ByVal outputCodePadrao As String, _
                                                 ByVal medidasAcessorios As Object) As String
    Dim texto As String
    Dim chave As Variant
    Dim prefixo As String
    Dim medida As String
    Dim qtdPorMedida As Long
    Dim qtdContabilizada As Long

    texto = ""
    prefixo = UCase$(nomeShape) & "_MEDIDA_"

    If Not medidasAcessorios Is Nothing Then
        For Each chave In medidasAcessorios.Keys
            If Left$(CStr(chave), Len(prefixo)) = prefixo Then
                qtdPorMedida = CLng(medidasAcessorios(chave))
                medida = Mid$(CStr(chave), Len(prefixo) + 1)

                texto = texto & MontarLinhaComQuantidade(qtdPorMedida, _
                    ResolverOutputCode(tipo, nomeShape, outputCodePadrao, medidasAcessorios, medida))
                qtdContabilizada = qtdContabilizada + qtdPorMedida
            End If
        Next chave

        If qtdContabilizada > 0 Then
            If qtdContabilizada < quantidadeTotal Then
                texto = texto & MontarLinhaComQuantidade((quantidadeTotal - qtdContabilizada), _
                        ResolverOutputCode(tipo, nomeShape, outputCodePadrao, medidasAcessorios))
            End If
            MontarLinhasAcessorioComMedida = texto
            Exit Function
        End If
    End If

    MontarLinhasAcessorioComMedida = MontarLinhaComQuantidade(quantidadeTotal, _
                                     ResolverOutputCode(tipo, nomeShape, outputCodePadrao, medidasAcessorios))
End Function

Private Function MontarLinhasKanbanPorGrupo(ByVal medidasAcessorios As Object) As String
    Dim texto As String
    Dim chave As Variant
    Dim assinatura As String
    Dim partes() As String
    Dim qtdGrupos As Long
    Dim qtdBaseNoGrupo As Long
    Dim qtdBasesTotal As Long
    Dim qtdTirasTotal As Long
    Dim qtdVD As Long
    Dim qtdAM As Long
    Dim qtdVM As Long
    Dim qtdCZ As Long
    Dim qtdPakIntNoGrupo As Long
    Dim qtdPakIntPorBase As Long
    Dim detalheCores As String
    Dim rotuloBase As String

    texto = ""
    If medidasAcessorios Is Nothing Then Exit Function

    For Each chave In medidasAcessorios.Keys
        If Left$(CStr(chave), 11) = "KANBAN_SIG_" Then
            assinatura = Mid$(CStr(chave), 12)
            partes = Split(assinatura, "|")

            If UBound(partes) >= 5 Then
                qtdGrupos = CLng(medidasAcessorios(chave))
                qtdBaseNoGrupo = CLng(partes(0))
                qtdBasesTotal = qtdBaseNoGrupo * qtdGrupos
                qtdTirasTotal = CLng(partes(1))
                qtdVD = CLng(partes(2))
                qtdAM = CLng(partes(3))
                qtdVM = CLng(partes(4))
                qtdCZ = CLng(partes(5))
                qtdPakIntNoGrupo = 0
                qtdPakIntPorBase = 0
                If UBound(partes) >= 6 Then
                    qtdPakIntNoGrupo = CLng(partes(6))
                    qtdPakIntPorBase = qtdPakIntNoGrupo
                    If qtdBaseNoGrupo > 0 Then
                        If qtdPakIntNoGrupo Mod qtdBaseNoGrupo = 0 Then
                            qtdPakIntPorBase = qtdPakIntNoGrupo \ qtdBaseNoGrupo
                        End If
                    End If
                End If

                rotuloBase = "BASE"
                If qtdBasesTotal > 1 Then rotuloBase = "BASES"

                detalheCores = MontarDetalheCoresKanban(qtdVD, qtdAM, qtdVM, qtdCZ)
                texto = texto & "- " & qtdBasesTotal & " " & rotuloBase & " KANBAN C/ " & qtdTirasTotal & _
                        " TIRAS T CADA" & vbCrLf & _
                        "(" & detalheCores & ")"
                If qtdPakIntPorBase > 0 Then
                    texto = texto & " + " & qtdPakIntPorBase & " PAK INT POR BASE"
                End If
                texto = texto & vbCrLf
            End If
        End If
    Next chave

    MontarLinhasKanbanPorGrupo = texto
End Function

Private Function MontarDetalheCoresKanban(ByVal qtdVD As Long, _
                                          ByVal qtdAM As Long, _
                                          ByVal qtdVM As Long, _
                                          ByVal qtdCZ As Long) As String
    Dim detalhe As String

    detalhe = ""
    AdicionarCorKanban detalhe, qtdVD, "VD"
    AdicionarCorKanban detalhe, qtdAM, "AM"
    AdicionarCorKanban detalhe, qtdVM, "VM"
    AdicionarCorKanban detalhe, qtdCZ, "CZ"

    MontarDetalheCoresKanban = detalhe
End Function

Private Sub AdicionarCorKanban(ByRef detalhe As String, _
                               ByVal quantidade As Long, _
                               ByVal cor As String)
    If quantidade <= 0 Then Exit Sub

    If detalhe <> "" Then detalhe = detalhe & ","
    detalhe = detalhe & quantidade & cor
End Sub

Private Function MontarLinhaComQuantidade(ByVal quantidade As Long, _
                                          ByVal descricao As String) As String
    MontarLinhaComQuantidade = "- " & quantidade & " " & descricao & vbCrLf
End Function

Private Function MontarLinhaSemQuantidade(ByVal descricao As String) As String
    MontarLinhaSemQuantidade = "- " & descricao & vbCrLf
End Function



