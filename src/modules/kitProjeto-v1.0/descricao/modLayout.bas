Attribute VB_Name = "modLayout"
' modLayout'
Option Explicit

Private Const TOLERANCIA_COR_CMYK As Double = 0.5
Private Const SHAPE_KSVR_A4_AD As String = "KSVR-A4-AD-MACRO"
Private Const SHAPE_KSVR_A4_MG As String = "KSVR-A4-MG-MACRO"
Private Const SHAPE_KSVP_A4_AD As String = "KSVP-A4-AD-MACRO"
Private Const SHAPE_KSVP_A4_MG As String = "KSVP-A4-MG-MACRO"
Private Const SHAPE_BASE_KANBAN As String = "BASE-KANBAN-MACRO"
Private Const SHAPE_TIRA_T_VD As String = "TIRA-T-VD-MACRO"
Private Const SHAPE_TIRA_T_AM As String = "TIRA-T-AM-MACRO"
Private Const SHAPE_TIRA_T_VM As String = "TIRA-T-VM-MACRO"
Private Const SHAPE_TIRA_T_CZ As String = "TIRA-T-CZ-MACRO"
Private Const SHAPE_PAK_INT As String = "PAK-INT-MACRO"

Public Function ObterRetanguloMagenta(ByRef retanguloBase As Shape) As Boolean
    Dim shapePagina As Shape
    Dim retanguloSelecionado As Shape
    Dim candidatos As Collection
    Dim maiorShape As Shape
    Dim maiorArea As Double
    Dim areaAtual As Double

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
        MsgBox "Nenhum retângulo magenta válido encontrado." & vbCrLf & _
               "Selecione manualmente o retângulo desejado e rode novamente.", vbExclamation
        ObterRetanguloMagenta = False
        Exit Function
    End If

    If candidatos.Count > 1 Then
        If ActiveSelectionRange.Count > 0 Then
            Set retanguloSelecionado = ActiveSelectionRange(1)
            If retanguloSelecionado.Type <> cdrRectangleShape Or _
               Not ShapeTemContornoMagenta(retanguloSelecionado) Then
                MsgBox "O retângulo selecionado não possui borda magenta válida." & vbCrLf & _
                       "Selecione manualmente um retângulo magenta válido e rode novamente.", vbExclamation
                ObterRetanguloMagenta = False
                Exit Function
            End If
            Set retanguloBase = retanguloSelecionado
        Else
            MsgBox "Mais de um retângulo magenta encontrado." & vbCrLf & _
                   "Selecione manualmente o retângulo desejado e rode novamente.", vbCritical
            ObterRetanguloMagenta = False
            Exit Function
        End If
    Else
        Set retanguloBase = maiorShape
    End If

    ObterRetanguloMagenta = True
End Function

Public Function TentarObterTextoSelecionado(ByRef textoSelecionado As Shape) As Boolean
    Dim sr As ShapeRange
    Dim sh As Shape

    TentarObterTextoSelecionado = False
    Set sr = ActiveSelectionRange

    For Each sh In sr
        If sh.Type = cdrTextShape Then
            Set textoSelecionado = sh
            TentarObterTextoSelecionado = True
            Exit Function
        End If
    Next sh
End Function

Public Function ColetarAcessorios(ByVal indice As Object, _
                                   ByRef ehMG As Boolean, _
                                   ByRef ehAD As Boolean, _
                                   ByRef medidasAcessorios As Object) As Object
    Dim contadores As Object
    Set contadores = InicializarContadores(indice)
    Set medidasAcessorios = CreateObject("Scripting.Dictionary")

    ehMG = False
    ehAD = False

    Dim sh As Shape
    For Each sh In ActivePage.Shapes
        ProcessarShape sh, indice, contadores, ehMG, ehAD, medidasAcessorios
    Next sh

    AplicarRegraReforcoAluminio contadores

    Set ColetarAcessorios = contadores
End Function

Private Function InicializarContadores(ByVal indice As Object) As Object
    Dim contadores As Object
    Dim chave As Variant

    Set contadores = CreateObject("Scripting.Dictionary")
    For Each chave In indice.Keys
        contadores.Add CStr(chave), 0
    Next chave

    Set InicializarContadores = contadores
End Function

Private Sub ProcessarShape(ByVal sh As Shape, _
                            ByVal indice As Object, _
                            ByRef contadores As Object, _
                            ByRef ehMG As Boolean, _
                            ByRef ehAD As Boolean, _
                            ByRef medidasAcessorios As Object)
    On Error GoTo ProximoShape

    Dim nomeShape As String
    Dim itemAcessorio As Object

    nomeShape = UCase$(sh.Name)
    RegistrarGrupoKanbanSeAplicavel sh, medidasAcessorios
    RegistrarVarianteBordaSeAplicavel nomeShape, sh, medidasAcessorios

    If indice.Exists(nomeShape) Then
        Dim medidaTexto As String

        contadores(nomeShape) = CLng(contadores(nomeShape)) + 1
        medidaTexto = FormatarMedidaTexto(sh.SizeHeight) & "x" & FormatarMedidaTexto(sh.SizeWidth)

        If Not medidasAcessorios.Exists(nomeShape) Then
            medidasAcessorios.Add nomeShape, medidaTexto
        End If
        IncrementarContadorMedidaAcessorio medidasAcessorios, nomeShape, medidaTexto

        Set itemAcessorio = indice(nomeShape)
        Select Case CStr(itemAcessorio("Compat"))
            Case COMPAT_MG: ehMG = True
            Case COMPAT_AD: ehAD = True
        End Select
    End If

    If sh.Type = cdrGroupShape Then
        Dim filho As Shape
        For Each filho In sh.Shapes
            ProcessarShape filho, indice, contadores, ehMG, ehAD, medidasAcessorios
        Next filho
    End If

    Exit Sub
ProximoShape:
    Err.Clear
End Sub

Private Sub RegistrarGrupoKanbanSeAplicavel(ByVal shapeGrupo As Shape, _
                                            ByRef medidasAcessorios As Object)
    If shapeGrupo.Type <> cdrGroupShape Then Exit Sub

    Dim qtdBase As Long
    Dim qtdVD As Long
    Dim qtdAM As Long
    Dim qtdVM As Long
    Dim qtdCZ As Long
    Dim qtdPakInt As Long
    Dim qtdTotalTiras As Long
    Dim chave As String

    qtdBase = ContarShapesPorNome(shapeGrupo, SHAPE_BASE_KANBAN)
    If qtdBase = 0 Then Exit Sub

    ContarTirasKanbanHibrido shapeGrupo, qtdVD, qtdAM, qtdVM, qtdCZ
    qtdPakInt = ContarShapesPorNome(shapeGrupo, SHAPE_PAK_INT)
    qtdTotalTiras = qtdVD + qtdAM + qtdVM + qtdCZ
    If qtdTotalTiras = 0 Then Exit Sub

    chave = "KANBAN_SIG_" & qtdBase & "|" & qtdTotalTiras & "|" & qtdVD & "|" & qtdAM & "|" & qtdVM & "|" & qtdCZ & "|" & qtdPakInt

    If medidasAcessorios.Exists(chave) Then
        medidasAcessorios(chave) = CLng(medidasAcessorios(chave)) + 1
    Else
        medidasAcessorios.Add chave, 1
    End If
End Sub

Private Sub ContarTirasKanbanHibrido(ByVal shapeGrupo As Shape, _
                                     ByRef qtdVD As Long, _
                                     ByRef qtdAM As Long, _
                                     ByRef qtdVM As Long, _
                                     ByRef qtdCZ As Long)
    ProcessarTirasKanbanRecursivo shapeGrupo, qtdVD, qtdAM, qtdVM, qtdCZ
End Sub

Private Sub ProcessarTirasKanbanRecursivo(ByVal sh As Shape, _
                                          ByRef qtdVD As Long, _
                                          ByRef qtdAM As Long, _
                                          ByRef qtdVM As Long, _
                                          ByRef qtdCZ As Long)
    On Error GoTo Fim

    If sh.Type = cdrGroupShape Then
        Dim filho As Shape
        For Each filho In sh.Shapes
            ProcessarTirasKanbanRecursivo filho, qtdVD, qtdAM, qtdVM, qtdCZ
        Next filho
        Exit Sub
    End If

    Dim nomeShape As String
    nomeShape = UCase$(sh.Name)
    If Not EhShapeTiraKanban(nomeShape) Then Exit Sub

    Dim siglaCor As String
    siglaCor = ObterSiglaCorTiraPorFill(sh)
    If siglaCor = "" Then Exit Sub

    IncrementarContadorCorKanban siglaCor, qtdVD, qtdAM, qtdVM, qtdCZ

Fim:
    Err.Clear
End Sub

Private Function EhShapeTiraKanban(ByVal nomeShape As String) As Boolean
    EhShapeTiraKanban = (Left$(UCase$(nomeShape), 7) = "TIRA-T-")
End Function

Private Function ObterSiglaCorTiraPorFill(ByVal sh As Shape) As String
    On Error GoTo Falha

    If sh.Fill Is Nothing Then Exit Function
    If sh.Fill.Type <> cdrUniformFill Then Exit Function

    Dim c As Double
    Dim m As Double
    Dim y As Double
    Dim k As Double

    c = sh.Fill.UniformColor.CMYKCyan
    m = sh.Fill.UniformColor.CMYKMagenta
    y = sh.Fill.UniformColor.CMYKYellow
    k = sh.Fill.UniformColor.CMYKBlack

    ' Tentativa 1: correspondencia exata com pequena tolerancia.
    If CorCMYKProxima(sh.Fill.UniformColor, 100, 0, 100, 0) Then
        ObterSiglaCorTiraPorFill = "VD"
        Exit Function
    End If

    If CorCMYKProxima(sh.Fill.UniformColor, 0, 0, 100, 0) Then
        ObterSiglaCorTiraPorFill = "AM"
        Exit Function
    End If

    If CorCMYKProxima(sh.Fill.UniformColor, 0, 100, 100, 0) Then
        ObterSiglaCorTiraPorFill = "VM"
        Exit Function
    End If

    If CorCMYKProxima(sh.Fill.UniformColor, 0, 0, 0, 30) Then
        ObterSiglaCorTiraPorFill = "CZ"
        Exit Function
    End If

    ' Tentativa 2: faixas tolerantes para lidar com conversoes de perfil de cor.
    ObterSiglaCorTiraPorFill = ClassificarCorTiraPorFaixa(c, m, y, k)

    Exit Function
Falha:
    ObterSiglaCorTiraPorFill = ""
    Err.Clear
End Function

Private Function ClassificarCorTiraPorFaixa(ByVal c As Double, _
                                            ByVal m As Double, _
                                            ByVal y As Double, _
                                            ByVal k As Double) As String
    ' Verde: C e Y altos, M baixo.
    If c >= 70 And y >= 70 And m <= 35 And k <= 35 Then
        ClassificarCorTiraPorFaixa = "VD"
        Exit Function
    End If

    ' Amarelo: Y alto com C/M baixos.
    If y >= 70 And c <= 25 And m <= 25 And k <= 25 Then
        ClassificarCorTiraPorFaixa = "AM"
        Exit Function
    End If

    ' Vermelho: M e Y altos com C baixo.
    If m >= 70 And y >= 70 And c <= 25 And k <= 25 Then
        ClassificarCorTiraPorFaixa = "VM"
        Exit Function
    End If

    ' Cinza: preto dominante com C/M/Y baixos.
    If k >= 20 And c <= 20 And m <= 20 And y <= 20 Then
        ClassificarCorTiraPorFaixa = "CZ"
    End If
End Function

Private Function CorCMYKProxima(ByVal cor As Color, _
                                ByVal c As Double, _
                                ByVal m As Double, _
                                ByVal y As Double, _
                                ByVal k As Double) As Boolean
    CorCMYKProxima = _
        Abs(cor.CMYKCyan - c) < TOLERANCIA_COR_CMYK And _
        Abs(cor.CMYKMagenta - m) < TOLERANCIA_COR_CMYK And _
        Abs(cor.CMYKYellow - y) < TOLERANCIA_COR_CMYK And _
        Abs(cor.CMYKBlack - k) < TOLERANCIA_COR_CMYK
End Function

Private Sub IncrementarContadorCorKanban(ByVal siglaCor As String, _
                                         ByRef qtdVD As Long, _
                                         ByRef qtdAM As Long, _
                                         ByRef qtdVM As Long, _
                                         ByRef qtdCZ As Long)
    Select Case UCase$(siglaCor)
        Case "VD"
            qtdVD = qtdVD + 1
        Case "AM"
            qtdAM = qtdAM + 1
        Case "VM"
            qtdVM = qtdVM + 1
        Case "CZ"
            qtdCZ = qtdCZ + 1
    End Select
End Sub

Private Sub IncrementarContadorMedidaAcessorio(ByRef medidasAcessorios As Object, _
                                               ByVal nomeShape As String, _
                                               ByVal medidaTexto As String)
    Dim chave As String
    chave = UCase$(nomeShape) & "_MEDIDA_" & medidaTexto

    If medidasAcessorios.Exists(chave) Then
        medidasAcessorios(chave) = CLng(medidasAcessorios(chave)) + 1
    Else
        medidasAcessorios.Add chave, 1
    End If
End Sub

Private Sub RegistrarVarianteBordaSeAplicavel(ByVal nomeShape As String, _
                                               ByVal shapeAcessorio As Shape, _
                                               ByRef medidasAcessorios As Object)
    If Not EhAcessorioComVarianteBorda(nomeShape) Then Exit Sub

    Dim variante As String
    variante = ObterVarianteBorda(shapeAcessorio)
    If variante = "" Then Exit Sub

    IncrementarContadorVariante medidasAcessorios, nomeShape, variante
End Sub

Private Function ObterVarianteBorda(ByVal shapeAcessorio As Shape) As String
    Select Case shapeAcessorio.Fill.Type
        Case cdrUniformFill
            ObterVarianteBorda = "UNIFORME"
        Case cdrFountainFill
            ObterVarianteBorda = "DEGRADÊ"
    End Select
End Function

Private Function EhAcessorioComVarianteBorda(ByVal nomeShape As String) As Boolean
    Select Case UCase$(nomeShape)
        Case SHAPE_KSVR_A4_AD, SHAPE_KSVR_A4_MG, SHAPE_KSVP_A4_AD, SHAPE_KSVP_A4_MG
            EhAcessorioComVarianteBorda = True
    End Select
End Function

Private Sub IncrementarContadorVariante(ByRef medidasAcessorios As Object, _
                                        ByVal nomeShape As String, _
                                        ByVal variante As String)
    Dim chave As String
    chave = ChaveQtdVariante(nomeShape, variante)

    If medidasAcessorios.Exists(chave) Then
        medidasAcessorios(chave) = CLng(medidasAcessorios(chave)) + 1
    Else
        medidasAcessorios.Add chave, 1
    End If
End Sub

Private Function ChaveQtdVariante(ByVal nomeShape As String, _
                                  ByVal variante As String) As String
    ChaveQtdVariante = UCase$(nomeShape) & "_VARIANTE_" & UCase$(variante) & "_QTD"
End Function

Private Sub AplicarRegraReforcoAluminio(ByRef contadores As Object)
    Const SHAPE_BASE_ESC_A4 As String = "BASE-ESC-A4-MACRO"
    Dim qtdBases As Long

    If contadores Is Nothing Then Exit Sub
    If Not contadores.Exists(SHAPE_BASE_ESC_A4) Then Exit Sub
    If Not contadores.Exists(SHAPE_REFORCO_ALUMINIO_AUTO) Then Exit Sub

    qtdBases = CLng(contadores(SHAPE_BASE_ESC_A4))
    If qtdBases < 6 Then Exit Sub

    contadores(SHAPE_REFORCO_ALUMINIO_AUTO) = 1
End Sub

Private Function FormatarMedidaTexto(ByVal medida As Double) As String
    Dim valorArredondado As Double
    valorArredondado = Round(medida, 1)

    If Abs(valorArredondado - CLng(valorArredondado)) < 0.0001 Then
        FormatarMedidaTexto = CStr(CLng(valorArredondado))
    Else
        FormatarMedidaTexto = Replace(Format$(valorArredondado, "0.0"), ".", ",")
    End If
End Function

Private Function ShapeTemContornoMagenta(ByVal s As Shape) As Boolean
    On Error GoTo Falha

    ShapeTemContornoMagenta = False
    If s Is Nothing Then Exit Function
    If s.Outline Is Nothing Then Exit Function

    ShapeTemContornoMagenta = _
        Abs(s.Outline.Color.CMYKCyan - 0) < TOLERANCIA_COR_CMYK And _
        Abs(s.Outline.Color.CMYKMagenta - 100) < TOLERANCIA_COR_CMYK And _
        Abs(s.Outline.Color.CMYKYellow - 0) < TOLERANCIA_COR_CMYK And _
        Abs(s.Outline.Color.CMYKBlack - 0) < TOLERANCIA_COR_CMYK

    Exit Function
Falha:
    ShapeTemContornoMagenta = False
    Err.Clear
End Function

Private Function ContarShapesPorNome(ByVal shapeRaiz As Shape, _
                                     ByVal nomeAlvo As String) As Long
    Dim filho As Shape

    If UCase$(shapeRaiz.Name) = UCase$(nomeAlvo) Then
        ContarShapesPorNome = ContarShapesPorNome + 1
    End If

    If shapeRaiz.Type <> cdrGroupShape Then Exit Function

    For Each filho In shapeRaiz.Shapes
        ContarShapesPorNome = ContarShapesPorNome + ContarShapesPorNome(filho, nomeAlvo)
    Next filho
End Function

