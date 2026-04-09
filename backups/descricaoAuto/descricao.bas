Attribute VB_Name = "descricaoAutov1"
Option Explicit

' Tipos de quadro aceitos no UserForm.
Private Const TIPO_QPMS_P As String = "QPMS-P"
Private Const TIPO_QPMM_P As String = "QPMM-P"

' Tolerância para validar magenta CMYK no contorno.
Private Const TOLERANCIA_COR_CMYK As Double = 0.5

' Prefixo padrão usado para identificar objetos de acessórios criados por macro.
Private Const SUFIXO_ACESSORIO_MACRO As String = "-MACRO"

Sub descricaoAuto()

    ' Objetos principais da rotina.
    Dim sh As Shape
    Dim rect As Shape
    Dim txt As Shape
    
    ' Medidas do quadro e texto final.
    Dim largura As Double
    Dim altura As Double
    Dim textoFinal As String
    
    ' Contadores de acessórios exibidos na descrição final.
    Dim qtdKSIR_AD As Long
    Dim qtdKSIP_AD As Long
    Dim qtdKSIR_MG As Long
    Dim qtdKSIP_MG As Long
    Dim qtdESC_A4_CZ As Long
    Dim qtdESC_A4_AM As Long
    Dim qtdESC_A4_AZ As Long
    Dim qtdESC_A4_VD As Long
    Dim qtdESC_A4_VM As Long
    Dim qtdESC_A4_PT As Long
    Dim qtdBASE_ESC_A4 As Long

    ' Texto consolidado da seção de acessórios.
    Dim textoAcessorios As String

    ' Flags de detecção genérica para validação de compatibilidade.
    Dim ehMagneticoDetectado As Boolean
    Dim ehAdesivoDetectado As Boolean
    
    ' Controle final do tipo do quadro selecionado no menu.
    Dim ehMagnetico As Boolean
    Dim codigoTipoQuadro As String

    ' 1) Solicita o tipo no UserForm.
    codigoTipoQuadro = SolicitarTipoQuadroUserForm()
    If codigoTipoQuadro = "" Then
        Exit Sub
    End If

    ' 2) Resolve automaticamente o retângulo magenta (com fallback híbrido).
    If Not TentarObterRetanguloMagenta(rect) Then
        Exit Sub
    End If

    ' 3) Exige que o texto de descrição esteja selecionado manualmente.
    If Not TentarObterTextoSelecionado(txt) Then
        MsgBox "Selecione manualmente o texto de descrição e rode novamente.", vbExclamation
        Exit Sub
    End If
    
    ' 4) Garante unidade em milímetros para leitura correta das medidas.
    ActiveDocument.Unit = cdrMillimeter
    
    ' 5) Captura dimensões do retângulo base.
    largura = Round(rect.SizeWidth, 0)
    altura = Round(rect.SizeHeight, 0)
    
    ' 6) Coleta acessórios na página e ativa flags de validação por padrão de nome.
    ColetarAcessoriosNaPagina qtdKSIR_AD, qtdKSIP_AD, qtdKSIR_MG, qtdKSIP_MG, _
                              qtdESC_A4_CZ, qtdESC_A4_AM, qtdESC_A4_AZ, qtdESC_A4_VD, _
                              qtdESC_A4_VM, qtdESC_A4_PT, qtdBASE_ESC_A4, _
                              ehMagneticoDetectado, ehAdesivoDetectado

    ' 7) Converte o tipo selecionado em regra binária de texto (magnético x normal).
    If Not DeterminarTipoMagnetico(codigoTipoQuadro, ehMagnetico) Then
        Exit Sub
    End If

    ' 8) Bloqueia inconsistência comum entre tipo escolhido e acessórios detectados.
    If Not ConfirmarCompatibilidadeAcessorios(codigoTipoQuadro, ehMagneticoDetectado, ehAdesivoDetectado) Then
        Exit Sub
    End If
    
    ' 9) Monta linhas de acessórios que serão anexadas ao texto principal.
    textoAcessorios = MontarTextoAcessorios(qtdKSIR_AD, qtdKSIP_AD, qtdKSIR_MG, qtdKSIP_MG, _
                                            qtdESC_A4_CZ, qtdESC_A4_AM, qtdESC_A4_AZ, qtdESC_A4_VD, _
                                            qtdESC_A4_VM, qtdESC_A4_PT, qtdBASE_ESC_A4)
    
    ' 10) Monta descrição principal do quadro.
    textoFinal = MontarTextoPrincipal(ehMagnetico, altura, largura)
    
    ' 11) Anexa bloco de acessórios somente quando existir conteúdo.
    textoFinal = AnexarSecaoAcessorios(textoFinal, textoAcessorios)
    
    ' 12) Atualiza o objeto de texto selecionado.
    txt.Text.Story = textoFinal
    
    ' 13) Feedback de conclusão.
    MsgBox "Texto atualizado com sucesso!", vbInformation

End Sub

' Converte o código escolhido no menu para regra magnético/não magnético.
Private Function DeterminarTipoMagnetico(ByVal codigoTipoQuadro As String, ByRef ehMagnetico As Boolean) As Boolean

    Select Case codigoTipoQuadro
        Case TIPO_QPMS_P
            ehMagnetico = False
        Case TIPO_QPMM_P
            ehMagnetico = True
        Case Else
            MsgBox "Tipo de quadro inválido: " & codigoTipoQuadro, vbCritical
            DeterminarTipoMagnetico = False
            Exit Function
    End Select

    DeterminarTipoMagnetico = True

End Function

' Lê todos os objetos da página e contabiliza os acessórios conhecidos.
Private Sub ColetarAcessoriosNaPagina(ByRef qtdKSIR_AD As Long, _
                                      ByRef qtdKSIP_AD As Long, _
                                      ByRef qtdKSIR_MG As Long, _
                                      ByRef qtdKSIP_MG As Long, _
                                      ByRef qtdESC_A4_CZ As Long, _
                                      ByRef qtdESC_A4_AM As Long, _
                                      ByRef qtdESC_A4_AZ As Long, _
                                      ByRef qtdESC_A4_VD As Long, _
                                      ByRef qtdESC_A4_VM As Long, _
                                      ByRef qtdESC_A4_PT As Long, _
                                      ByRef qtdBASE_ESC_A4 As Long, _
                                      ByRef ehMagneticoDetectado As Boolean, _
                                      ByRef ehAdesivoDetectado As Boolean)

    Dim sh As Shape

    InicializarContadoresAcessorios qtdKSIR_AD, qtdKSIP_AD, qtdKSIR_MG, qtdKSIP_MG, _
                                    qtdESC_A4_CZ, qtdESC_A4_AM, qtdESC_A4_AZ, qtdESC_A4_VD, _
                                    qtdESC_A4_VM, qtdESC_A4_PT, qtdBASE_ESC_A4

    ehMagneticoDetectado = False
    ehAdesivoDetectado = False

    For Each sh In ActivePage.Shapes
        ProcessarShapeAcessorios sh, qtdKSIR_AD, qtdKSIP_AD, qtdKSIR_MG, qtdKSIP_MG, _
                                qtdESC_A4_CZ, qtdESC_A4_AM, qtdESC_A4_AZ, qtdESC_A4_VD, _
                                qtdESC_A4_VM, qtdESC_A4_PT, qtdBASE_ESC_A4, ehMagneticoDetectado, ehAdesivoDetectado
    Next sh

End Sub

' Anexa seção formatada de acessórios ao texto principal quando houver itens.
Private Function AnexarSecaoAcessorios(ByVal textoPrincipal As String, ByVal textoAcessorios As String) As String

    If textoAcessorios = "" Then
        AnexarSecaoAcessorios = textoPrincipal
    Else
        AnexarSecaoAcessorios = textoPrincipal & vbCrLf & vbCrLf & _
                                "ACESSÓRIOS:" & vbCrLf & vbCrLf & _
                                textoAcessorios
    End If

End Function

Private Function ConfirmarCompatibilidadeAcessorios(ByVal codigoTipoQuadro As String, _
                                                    ByVal encontrouAcessorioMG As Boolean, _
                                                    ByVal encontrouAcessorioAD As Boolean) As Boolean

    ' Resultado do diálogo Sim/Não com o usuário.
    Dim resposta As VbMsgBoxResult

    ' Proteção: QPMS normalmente não deve conter acessórios magnéticos.
    If codigoTipoQuadro = "QPMS-P" And encontrouAcessorioMG Then
        resposta = MsgBox("Acessórios magnéticos encontrados, deseja continuar?", vbQuestion + vbYesNo, "Validação de acessórios")
        If resposta = vbNo Then
            ConfirmarCompatibilidadeAcessorios = False
            Exit Function
        End If
    End If

    ' Proteção: QPMM normalmente não deve conter acessórios adesivos.
    If codigoTipoQuadro = "QPMM-P" And encontrouAcessorioAD Then
        resposta = MsgBox("Acessórios adesivos encontrados, deseja continuar?", vbQuestion + vbYesNo, "Validação de acessórios")
        If resposta = vbNo Then
            ConfirmarCompatibilidadeAcessorios = False
            Exit Function
        End If
    End If

    ConfirmarCompatibilidadeAcessorios = True

End Function

' Localiza um texto na seleção atual para ser atualizado pela macro.
Private Function TentarObterTextoSelecionado(ByRef textoSelecionado As Shape) As Boolean

    ' Faixa de objetos atualmente selecionados.
    Dim sr As ShapeRange

    ' Iterador da seleção.
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

' Zera todos os contadores de acessórios antes da leitura da página.
Private Sub InicializarContadoresAcessorios(ByRef qtdKSIR_AD As Long, _
                                            ByRef qtdKSIP_AD As Long, _
                                            ByRef qtdKSIR_MG As Long, _
                                            ByRef qtdKSIP_MG As Long, _
                                            ByRef qtdESC_A4_CZ As Long, _
                                            ByRef qtdESC_A4_AM As Long, _
                                            ByRef qtdESC_A4_AZ As Long, _
                                            ByRef qtdESC_A4_VD As Long, _
                                            ByRef qtdESC_A4_VM As Long, _
                                            ByRef qtdESC_A4_PT As Long, _
                                            ByRef qtdBASE_ESC_A4 As Long)

    qtdKSIR_AD = 0
    qtdKSIP_AD = 0
    qtdKSIR_MG = 0
    qtdKSIP_MG = 0
    qtdESC_A4_CZ = 0
    qtdESC_A4_AM = 0
    qtdESC_A4_AZ = 0
    qtdESC_A4_VD = 0
    qtdESC_A4_VM = 0
    qtdESC_A4_PT = 0
    qtdBASE_ESC_A4 = 0

End Sub

' Constrói apenas as linhas de acessórios com quantidade maior que zero.
Private Function MontarTextoAcessorios(ByVal qtdKSIR_AD As Long, _
                                       ByVal qtdKSIP_AD As Long, _
                                       ByVal qtdKSIR_MG As Long, _
                                       ByVal qtdKSIP_MG As Long, _
                                       ByVal qtdESC_A4_CZ As Long, _
                                       ByVal qtdESC_A4_AM As Long, _
                                       ByVal qtdESC_A4_AZ As Long, _
                                       ByVal qtdESC_A4_VD As Long, _
                                       ByVal qtdESC_A4_VM As Long, _
                                       ByVal qtdESC_A4_PT As Long, _
                                       ByVal qtdBASE_ESC_A4 As Long) As String

    Dim texto As String
    texto = ""

    If qtdKSIR_AD > 0 Then texto = texto & "- " & qtdKSIR_AD & " KSIR-A4-AD" & vbCrLf
    If qtdKSIP_AD > 0 Then texto = texto & "- " & qtdKSIP_AD & " KSIP-A4-AD" & vbCrLf
    If qtdKSIR_MG > 0 Then texto = texto & "- " & qtdKSIR_MG & " KSIR-A4-MG" & vbCrLf
    If qtdKSIP_MG > 0 Then texto = texto & "- " & qtdKSIP_MG & " KSIP-A4-MG" & vbCrLf
    If qtdESC_A4_CZ > 0 Then texto = texto & "- " & qtdESC_A4_CZ & " ESC-A4-CZ" & vbCrLf
    If qtdESC_A4_AM > 0 Then texto = texto & "- " & qtdESC_A4_AM & " ESC-A4-AM" & vbCrLf
    If qtdESC_A4_AZ > 0 Then texto = texto & "- " & qtdESC_A4_AZ & " ESC-A4-AZ" & vbCrLf
    If qtdESC_A4_VD > 0 Then texto = texto & "- " & qtdESC_A4_VD & " ESC-A4-VD" & vbCrLf
    If qtdESC_A4_VM > 0 Then texto = texto & "- " & qtdESC_A4_VM & " ESC-A4-VM" & vbCrLf
    If qtdESC_A4_PT > 0 Then texto = texto & "- " & qtdESC_A4_PT & " ESC-A4-PT" & vbCrLf
    If qtdBASE_ESC_A4 > 0 Then texto = texto & "- " & qtdBASE_ESC_A4 & " BASE-ESC-A4" & vbCrLf

    MontarTextoAcessorios = texto

End Function

' Monta a descrição base do quadro de acordo com o tipo final definido.
Private Function MontarTextoPrincipal(ByVal ehMagnetico As Boolean, ByVal altura As Double, ByVal largura As Double) As String

    If ehMagnetico Then
        MontarTextoPrincipal = "QUADRO BRANCO MAGNÉTICO" & vbCrLf & _
                              "PARA ESCRITA COM IMPRESSÃO " & vbCrLf & _
                              "DIGITAL UV. E LAMINAÇÃO PYT" & vbCrLf & _
                              "MED " & altura & "x" & largura & " - QPMM"
    Else
        MontarTextoPrincipal = "QUADRO BRANCO PARA ESCRITA" & vbCrLf & _
                              "COM IMPRESSÃO DIGITAL UV. E" & vbCrLf & _
                              "LAMINAÇÃO PYT MED " & altura & "x" & largura & vbCrLf & _
                              "- QPMS"
    End If

End Function

' Abre o UserForm de seleção e retorna o tipo escolhido no ComboBox.
Private Function SolicitarTipoQuadroUserForm() As String

    On Error GoTo FalhaUserForm

    frmTipoQuadro.TipoSelecionado = ""
    frmTipoQuadro.Show vbModal

    SolicitarTipoQuadroUserForm = frmTipoQuadro.TipoSelecionado
    Unload frmTipoQuadro
    Exit Function

FalhaUserForm:
    MsgBox "UserForm 'frmTipoQuadro' não encontrado ou inválido." & vbCrLf & _
        "Crie o formulário com ComboBox conforme o passo a passo.", vbCritical
    SolicitarTipoQuadroUserForm = ""

End Function

' Resolve o retângulo magenta usando lógica híbrida (automática + seleção manual).
Private Function TentarObterRetanguloMagenta(ByRef retanguloBase As Shape) As Boolean

    ' Iterador de shapes da página.
    Dim shapePagina As Shape

    ' Seleção manual usada quando há mais de um candidato magenta.
    Dim retanguloSelecionado As Shape

    ' Coleção de candidatos magenta válidos.
    Dim candidatos As Collection

    ' Candidato de maior área (escolha automática padrão).
    Dim maiorShape As Shape

    ' Controle de comparação de área.
    Dim maiorArea As Double
    Dim areaAtual As Double

    Set candidatos = New Collection
    maiorArea = 0

    For Each shapePagina In ActivePage.Shapes
        If shapePagina.Type = cdrRectangleShape Then
            If ShapeTemContornoMagentaDesc(shapePagina) Then
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
        MsgBox "Nenhum retângulo magenta válido foi encontrado automaticamente." & vbCrLf & _
               "Selecione manualmente o retângulo desejado e rode novamente.", vbExclamation
        TentarObterRetanguloMagenta = False
        Exit Function
    End If

    If candidatos.Count > 1 Then
        If ActiveSelectionRange.Count > 0 Then
            Set retanguloSelecionado = ActiveSelectionRange(1)

            If retanguloSelecionado.Type <> cdrRectangleShape Then
                MsgBox "Mais de um retângulo magenta encontrado." & vbCrLf & _
                       "Selecione manualmente o retângulo desejado e rode novamente.", vbExclamation
                TentarObterRetanguloMagenta = False
                Exit Function
            End If

            If Not ShapeTemContornoMagentaDesc(retanguloSelecionado) Then
                MsgBox "O retângulo selecionado não possui borda magenta CMYK válida." & vbCrLf & _
                       "Selecione manualmente um retângulo magenta válido e rode novamente.", vbExclamation
                TentarObterRetanguloMagenta = False
                Exit Function
            End If

            Set retanguloBase = retanguloSelecionado
        Else
            MsgBox "Mais de um retângulo magenta encontrado." & vbCrLf & _
                   "Selecione manualmente o retângulo desejado e rode novamente.", vbCritical
            TentarObterRetanguloMagenta = False
            Exit Function
        End If
    Else
        Set retanguloBase = maiorShape
    End If

    TentarObterRetanguloMagenta = True

End Function

' Valida se uma shape possui contorno magenta em CMYK dentro da tolerância.
Private Function ShapeTemContornoMagentaDesc(ByVal s As Shape) As Boolean
    On Error GoTo Falha

    ShapeTemContornoMagentaDesc = False

    If s Is Nothing Then Exit Function
    If s.Outline Is Nothing Then Exit Function

    ShapeTemContornoMagentaDesc = _
        Abs(s.Outline.Color.CMYKCyan - 0) < TOLERANCIA_COR_CMYK And _
        Abs(s.Outline.Color.CMYKMagenta - 100) < TOLERANCIA_COR_CMYK And _
        Abs(s.Outline.Color.CMYKYellow - 0) < TOLERANCIA_COR_CMYK And _
        Abs(s.Outline.Color.CMYKBlack - 0) < TOLERANCIA_COR_CMYK

    Exit Function

Falha:
    ShapeTemContornoMagentaDesc = False
    Err.Clear
End Function

' Conta acessórios conhecidos e detecta padrões AD/MG para validação de tipo.
Private Sub ProcessarShapeAcessorios(ByVal sh As Shape, _
                                    ByRef qtdKSIR_AD As Long, _
                                    ByRef qtdKSIP_AD As Long, _
                                    ByRef qtdKSIR_MG As Long, _
                                    ByRef qtdKSIP_MG As Long, _
                                    ByRef qtdESC_A4_CZ As Long, _
                                    ByRef qtdESC_A4_AM As Long, _
                                    ByRef qtdESC_A4_AZ As Long, _
                                    ByRef qtdESC_A4_VD As Long, _
                                    ByRef qtdESC_A4_VM As Long, _
                                    ByRef qtdESC_A4_PT As Long, _
                                    ByRef qtdBASE_ESC_A4 As Long, _
                                    ByRef ehMagnetico As Boolean, _
                                    ByRef ehAdesivo As Boolean)

    ' Iterador de filhos quando a shape é um grupo.
    Dim filho As Shape

    ' Detecta padrão de acessórios por nome para validar compatibilidade futura.
    If InStr(1, sh.Name, SUFIXO_ACESSORIO_MACRO, vbTextCompare) > 0 Then
        If InStr(1, sh.Name, "-MG", vbTextCompare) > 0 Then
            ehMagnetico = True
        End If
        If InStr(1, sh.Name, "-AD", vbTextCompare) > 0 Then
            ehAdesivo = True
        End If
    End If

    Select Case sh.Name
        Case "KSIR-A4-AD-MACRO"
            qtdKSIR_AD = qtdKSIR_AD + 1
        Case "KSIP-A4-AD-MACRO"
            qtdKSIP_AD = qtdKSIP_AD + 1
        Case "KSIR-A4-MG-MACRO"
            qtdKSIR_MG = qtdKSIR_MG + 1
        Case "KSIP-A4-MG-MACRO"
            qtdKSIP_MG = qtdKSIP_MG + 1
        Case "ESC-A4-CZ-MACRO"
            qtdESC_A4_CZ = qtdESC_A4_CZ + 1
        Case "ESC-A4-AM-MACRO"
            qtdESC_A4_AM = qtdESC_A4_AM + 1
        Case "ESC-A4-AZ-MACRO"
            qtdESC_A4_AZ = qtdESC_A4_AZ + 1
        Case "ESC-A4-VD-MACRO"
            qtdESC_A4_VD = qtdESC_A4_VD + 1
        Case "ESC-A4-VM-MACRO"
            qtdESC_A4_VM = qtdESC_A4_VM + 1
        Case "ESC-A4-PT-MACRO"
            qtdESC_A4_PT = qtdESC_A4_PT + 1
        Case "BASE-ESC-A4-MACRO"
            qtdBASE_ESC_A4 = qtdBASE_ESC_A4 + 1
    End Select

    If sh.Type = cdrGroupShape Then
        For Each filho In sh.Shapes
            ProcessarShapeAcessorios filho, qtdKSIR_AD, qtdKSIP_AD, qtdKSIR_MG, qtdKSIP_MG, _
                                    qtdESC_A4_CZ, qtdESC_A4_AM, qtdESC_A4_AZ, qtdESC_A4_VD, _
                                    qtdESC_A4_VM, qtdESC_A4_PT, qtdBASE_ESC_A4, ehMagnetico, ehAdesivo
        Next filho
    End If

End Sub

