Attribute VB_Name = "Val"
Option Explicit

' =========================================================
' VALIDAÇÕES E IMPORTAÇÃO
' =========================================================

Public Function ObterQuadroMagentaValido() As Shape
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

Public Function ArquivoExiste(ByVal caminhoArquivo As String) As Boolean
    ArquivoExiste = (Dir$(caminhoArquivo) <> "")
End Function

Public Function ImportarArquivoCavalete(ByVal caminhoArquivo As String) As Shape
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


