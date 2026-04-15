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
            outputCode = ResolverOutputCode(nomeShape, CStr(item("OutputCode")), medidasAcessorios)
            If EhAcessorioComVarianteBorda(nomeShape) Then
                texto = texto & MontarLinhasAcessorioComVariante(nomeShape, quantidade, outputCode, medidasAcessorios)
            ElseIf EhCavaleteMetalon3(nomeShape) Then
                texto = texto & "- " & outputCode & vbCrLf
            Else
                texto = texto & "- " & quantidade & " " & outputCode & vbCrLf
            End If
        End If
    Next item

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
    If UCase$(nomeShape) = "TESTEIRA-MACRO" Then
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

