Attribute VB_Name = "modDescricao"
' modDescricao
Option Explicit

Private Const SHAPE_KSVR_A4_AD As String = "KSVR-A4-AD-MACRO"
Private Const SHAPE_KSVR_A4_MG As String = "KSVR-A4-MG-MACRO"
Private Const SHAPE_KSVP_A4_AD As String = "KSVP-A4-AD-MACRO"
Private Const SHAPE_KSVP_A4_MG As String = "KSVP-A4-MG-MACRO"

Public Function MontarTextoCompleto(ByVal tipo As tipoQuadro, _
                                     ByVal altura As Double, _
                                     ByVal largura As Double, _
                                     ByVal catalogo As Collection, _
                                     ByVal contadores As Object, _
                                     ByVal medidasAcessorios As Object) As String
    Dim texto As String
    texto = MontarTextoPrincipal(tipo, altura, largura)

    Dim secaoAcessorios As String
    secaoAcessorios = MontarTextoAcessorios(catalogo, contadores, medidasAcessorios)

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

Private Function MontarTextoAcessorios(ByVal catalogo As Collection, _
                                        ByVal contadores As Object, _
                                        ByVal medidasAcessorios As Object) As String
    Dim item As Variant
    Dim nomeShape As String
    Dim quantidade As Long
    Dim texto As String
    Dim outputCode As String

    texto = ""

    For Each item In catalogo
        nomeShape = CStr(item("ShapeName"))
        quantidade = CLng(contadores(nomeShape))
        If quantidade > 0 Then
            outputCode = CStr(item("OutputCode"))
            If EhAcessorioComMedidaSeparada(nomeShape) Then
                texto = texto & MontarLinhasAcessorioComMedida(nomeShape, quantidade, outputCode, medidasAcessorios)
            ElseIf EhAcessorioComVarianteBorda(nomeShape) Then
                outputCode = ResolverOutputCode(nomeShape, outputCode, medidasAcessorios)
                texto = texto & MontarLinhasAcessorioComVariante(nomeShape, quantidade, outputCode, medidasAcessorios)
            ElseIf EhCavaleteMetalon3(nomeShape) Then
                texto = texto & "- " & outputCode & vbCrLf
            Else
                outputCode = ResolverOutputCode(nomeShape, outputCode, medidasAcessorios)
                texto = texto & "- " & quantidade & " " & outputCode & vbCrLf
            End If
        End If
    Next item

    texto = texto & MontarLinhasKanbanPorGrupo(medidasAcessorios)

    MontarTextoAcessorios = texto
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

Private Function ResolverOutputCode(ByVal nomeShape As String, _
                                    ByVal outputCode As String, _
                                    ByVal medidasAcessorios As Object) As String
    If UCase$(nomeShape) = "TESTEIRA-MACRO" Or UCase$(nomeShape) = "DAVN-MACRO" Then
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

Private Function ChaveVariantePorShape(ByVal nomeShape As String) As String
    ChaveVariantePorShape = UCase$(nomeShape) & "_VARIANTE"
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
        Case "TESTEIRA-MACRO", "DAVN-MACRO"
            EhAcessorioComMedidaSeparada = True
    End Select
End Function

Private Function MontarLinhasAcessorioComMedida(ByVal nomeShape As String, _
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

                texto = texto & "- " & qtdPorMedida & " " & _
                        Replace(outputCodePadrao, "ALTXLARGURA", medida) & vbCrLf
                qtdContabilizada = qtdContabilizada + qtdPorMedida
            End If
        Next chave

        If qtdContabilizada > 0 Then
            If qtdContabilizada < quantidadeTotal Then
                texto = texto & "- " & (quantidadeTotal - qtdContabilizada) & " " & _
                        ResolverOutputCode(nomeShape, outputCodePadrao, medidasAcessorios) & vbCrLf
            End If
            MontarLinhasAcessorioComMedida = texto
            Exit Function
        End If
    End If

    MontarLinhasAcessorioComMedida = "- " & quantidadeTotal & " " & _
                                     ResolverOutputCode(nomeShape, outputCodePadrao, medidasAcessorios) & vbCrLf
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

                detalheCores = MontarDetalheCoresKanban(qtdVD, qtdAM, qtdVM, qtdCZ)
                texto = texto & "- " & qtdBasesTotal & " BASE KANBAN C/ " & qtdTirasTotal & _
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
