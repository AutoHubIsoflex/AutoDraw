Attribute VB_Name = "Módulo1"
Option Explicit

Private Const CAMINHO_ESCANINHO As String = "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\ESCANINHOS\ISOLEAN A4\20 ESC MACROCASCATA.cdr"
Private Const CAMINHO_ESTRUTURA As String = "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\CASCATA\ESTRUTURA.cdr"

Private Const MAX_ESCANINHOS_DESENHO As Long = 20
Private Const LIMITE_COLUNAS_POR_QUADRO As Long = 8
Private Const LIMITE_ESCANINHOS_POR_COLUNA As Long = 20
Private Const DESLOCAMENTO_ESTRUTURA_BAIXO As Double = 108
Private Const SOBRA_INT_LATERAL As Double = 30

Private Const NOME_BASE_ALINHAMENTO As String = "BASE-ESC-A4-MACRO"
Private Const NOME_GABARITO As String = "GABARITO"

Private Const NOME_INT1_CASC As String = "INT1-CASC"
Private Const NOME_INT2_CASC As String = "INT2-CASC"
Private Const NOME_INT3_CASC As String = "INT3-CASC"

Private Const NOME_INTSUP_ESTRUTURA As String = "INTSUP-ESTRT"
Private Const NOME_INTDIR_ESTRUTURA As String = "INTDIR-ESTRT"
Private Const NOME_INTESQ_ESTRUTURA As String = "INTESQ-ESTRT"

Sub Gerar_Escaninhos_Importando_CDR()

    Dim txtColunas As String
    Dim txtEscaninhos As String
    Dim qtdColunas As Long
    Dim escPorColuna() As Long
    Dim partes() As String
    Dim i As Long

    Dim coluna As Shape
    Dim grupoEscaninhos As Shape
    Dim estrutura As Shape
    Dim colunasImportadas As ShapeRange
    Dim esquerdaAtual As Double
    Dim topoBase As Double

    If Documents.Count = 0 Then
        MsgBox "Abra ou crie um documento no CorelDRAW antes de executar a macro.", vbExclamation
        Exit Sub
    End If

    If Dir$(CAMINHO_ESCANINHO) = "" Then
        MsgBox "Arquivo nao encontrado:" & vbCrLf & CAMINHO_ESCANINHO, vbCritical
        Exit Sub
    End If

    If Dir$(CAMINHO_ESTRUTURA) = "" Then
        MsgBox "Arquivo nao encontrado:" & vbCrLf & CAMINHO_ESTRUTURA, vbCritical
        Exit Sub
    End If

    ActiveDocument.Unit = cdrMillimeter

    txtColunas = InputBox("Quantidade de colunas?", "Gerar escaninhos")
    If Trim$(txtColunas) = "" Then Exit Sub

    If Not EhInteiroPositivo(txtColunas) Then
        MsgBox "Informe uma quantidade de colunas valida.", vbExclamation
        Exit Sub
    End If

    qtdColunas = CLng(txtColunas)

    If qtdColunas > LIMITE_COLUNAS_POR_QUADRO Then
        MsgBox "Cada quadro permite no maximo " & LIMITE_COLUNAS_POR_QUADRO & " colunas.", vbExclamation
        Exit Sub
    End If

    txtEscaninhos = InputBox( _
        "Qnt esc p/ coluna?" & vbCrLf & vbCrLf & _
        "Digite um numero unico para repetir em todas as colunas." & vbCrLf & _
        "Ou digite uma lista separada por virgula." & vbCrLf & _
        "Exemplo: 4 ou 3,4,5,4,3", _
        "Gerar escaninhos" _
    )

    If Trim$(txtEscaninhos) = "" Then Exit Sub

    txtEscaninhos = Replace(txtEscaninhos, " ", "")
    txtEscaninhos = Replace(txtEscaninhos, ";", ",")

    partes = Split(txtEscaninhos, ",")

    ReDim escPorColuna(1 To qtdColunas)

    If UBound(partes) = 0 Then

        If Not EhInteiroPositivo(partes(0)) Then
            MsgBox "Informe uma quantidade de escaninhos valida.", vbExclamation
            Exit Sub
        End If

        If CLng(partes(0)) > LIMITE_ESCANINHOS_POR_COLUNA Then
            MsgBox "Cada coluna permite no maximo " & LIMITE_ESCANINHOS_POR_COLUNA & " escaninhos.", vbExclamation
            Exit Sub
        End If

        For i = 1 To qtdColunas
            escPorColuna(i) = CLng(partes(0))
        Next i

    Else

        If UBound(partes) + 1 <> qtdColunas Then
            MsgBox "A quantidade informada nao bate com o numero de colunas." & vbCrLf & _
                   "Exemplo: para 5 colunas, use algo como 3,4,5,4,3.", vbExclamation
            Exit Sub
        End If

        For i = 1 To qtdColunas
            If Not EhInteiroPositivo(partes(i - 1)) Then
                MsgBox "Valor invalido na coluna " & i & ".", vbExclamation
                Exit Sub
            End If

            If CLng(partes(i - 1)) > LIMITE_ESCANINHOS_POR_COLUNA Then
                MsgBox "A coluna " & i & " passou de " & LIMITE_ESCANINHOS_POR_COLUNA & " escaninhos.", vbExclamation
                Exit Sub
            End If

            escPorColuna(i) = CLng(partes(i - 1))
        Next i

    End If

    ActiveDocument.BeginCommandGroup "Gerar escaninhos importados"

    On Error GoTo TrataErro

    Set colunasImportadas = New ShapeRange

    esquerdaAtual = 20
    topoBase = ActivePage.SizeHeight - 20

    For i = 1 To qtdColunas

        Set coluna = ImportarColunaEscaninho()

        coluna.Name = "COLUNA-A4-" & Format$(i, "00")

        AjustarQuantidadeEscaninhos coluna, escPorColuna(i)

        PosicionarPeloTopoEsquerdo coluna, esquerdaAtual, topoBase

        esquerdaAtual = ObterDireita(coluna)

        colunasImportadas.Add coluna

    Next i

    RemoverObjetosPorNome NOME_GABARITO

    Set grupoEscaninhos = AgruparShapeRange(colunasImportadas, "GRUPO-ESCANINHOS-A4-MACRO")

    Set estrutura = ImportarEstrutura()
    estrutura.Name = "ESTRUTURA-CASCATA-MACRO"

    AlinharEstruturaNaBase estrutura, grupoEscaninhos

    estrutura.Move 0, -DESLOCAMENTO_ESTRUTURA_BAIXO

    AjustarIntermediariosCascata estrutura, grupoEscaninhos, escPorColuna

    AjustarLarguraObjetosInt estrutura, grupoEscaninhos

    grupoEscaninhos.OrderToFront

    ActiveDocument.EndCommandGroup

    MsgBox "Escaninhos gerados com sucesso.", vbInformation
    Exit Sub

TrataErro:
    ActiveDocument.EndCommandGroup
    MsgBox "Erro ao gerar escaninhos:" & vbCrLf & Err.Description, vbCritical

End Sub

Private Function ImportarColunaEscaninho() As Shape
    Set ImportarColunaEscaninho = ImportarArquivoCDR(CAMINHO_ESCANINHO, "escaninho")
End Function

Private Function ImportarEstrutura() As Shape
    Set ImportarEstrutura = ImportarArquivoCDR(CAMINHO_ESTRUTURA, "estrutura")
End Function

Private Function ImportarArquivoCDR(ByVal caminhoArquivo As String, ByVal descricaoArquivo As String) As Shape

    Dim filtro As ImportFilter
    Dim sr As ShapeRange

    ActiveDocument.ClearSelection

    Set filtro = ActiveLayer.ImportEx(caminhoArquivo, cdrCDR)
    filtro.Finish

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        Err.Raise vbObjectError + 100, , "Nada foi importado do arquivo de " & descricaoArquivo & "."
    End If

    If sr.Count = 1 Then
        Set ImportarArquivoCDR = sr(1)
    Else
        Set ImportarArquivoCDR = sr.Group
    End If

End Function

Private Sub AjustarQuantidadeEscaninhos(ByVal coluna As Shape, ByVal qtdManter As Long)

    Dim escaninhos() As Shape
    Dim total As Long
    Dim i As Long

    ReDim escaninhos(1 To MAX_ESCANINHOS_DESENHO * 2)

    ColetarEscaninhos coluna, escaninhos, total

    If total = 0 Then
        Err.Raise vbObjectError + 101, , "Nenhum objeto com nome iniciado por ESC-A4 foi encontrado."
    End If

    If qtdManter > total Then
        Err.Raise vbObjectError + 102, , "O desenho importado possui apenas " & total & " escaninhos ESC-A4."
    End If

    OrdenarEscaninhos escaninhos, total

    For i = 1 To total - qtdManter
        escaninhos(i).Delete
    Next i

End Sub

Private Sub ColetarEscaninhos(ByVal shp As Shape, ByRef escaninhos() As Shape, ByRef total As Long)

    Dim filho As Shape

    If UCase$(Left$(shp.Name, 6)) = "ESC-A4" Then
        total = total + 1

        If total > UBound(escaninhos) Then
            ReDim Preserve escaninhos(1 To total + 25)
        End If

        Set escaninhos(total) = shp
    End If

    If shp.Type = cdrGroupShape Then
        For Each filho In shp.Shapes
            ColetarEscaninhos filho, escaninhos, total
        Next filho
    End If

End Sub

Private Sub OrdenarEscaninhos(ByRef escaninhos() As Shape, ByVal total As Long)

    Dim i As Long
    Dim j As Long
    Dim usaNumeroNome As Boolean
    Dim chaveI As Double
    Dim chaveJ As Double
    Dim temp As Shape

    usaNumeroNome = TodosTemNumeroNoNome(escaninhos, total)

    For i = 1 To total - 1
        For j = i + 1 To total

            chaveI = ChaveEscaninho(escaninhos(i), usaNumeroNome)
            chaveJ = ChaveEscaninho(escaninhos(j), usaNumeroNome)

            If chaveI > chaveJ Then
                Set temp = escaninhos(i)
                Set escaninhos(i) = escaninhos(j)
                Set escaninhos(j) = temp
            End If

        Next j
    Next i

End Sub

Private Function TodosTemNumeroNoNome(ByRef escaninhos() As Shape, ByVal total As Long) As Boolean

    Dim i As Long

    TodosTemNumeroNoNome = True

    For i = 1 To total
        If UltimoNumeroNoTexto(escaninhos(i).Name) = 0 Then
            TodosTemNumeroNoNome = False
            Exit Function
        End If
    Next i

End Function

Private Function ChaveEscaninho(ByVal shp As Shape, ByVal usaNumeroNome As Boolean) As Double

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    If usaNumeroNome Then
        ChaveEscaninho = UltimoNumeroNoTexto(shp.Name)
    Else
        shp.GetBoundingBox x, y, w, h
        ChaveEscaninho = -(y + (h / 2))
    End If

End Function

Private Function UltimoNumeroNoTexto(ByVal texto As String) As Long

    Dim i As Long
    Dim numero As String
    Dim caractere As String

    For i = Len(texto) To 1 Step -1
        caractere = Mid$(texto, i, 1)

        If caractere >= "0" And caractere <= "9" Then
            numero = caractere & numero
        ElseIf Len(numero) > 0 Then
            Exit For
        End If
    Next i

    If Len(numero) > 0 Then
        UltimoNumeroNoTexto = CLng(numero)
    Else
        UltimoNumeroNoTexto = 0
    End If

End Function

Private Function AgruparShapeRange(ByVal sr As ShapeRange, ByVal nomeGrupo As String) As Shape

    If sr.Count = 0 Then
        Err.Raise vbObjectError + 103, , "Nenhuma coluna foi encontrada para agrupar."
    End If

    If sr.Count = 1 Then
        Set AgruparShapeRange = sr(1)
    Else
        Set AgruparShapeRange = sr.Group
    End If

    AgruparShapeRange.Name = nomeGrupo

End Function

Private Sub AlinharEstruturaNaBase(ByVal estrutura As Shape, ByVal grupoEscaninhos As Shape)

    Dim bases As ShapeRange

    Set bases = New ShapeRange

    ColetarObjetosPorNomeNoShape grupoEscaninhos, NOME_BASE_ALINHAMENTO, bases

    If bases.Count = 0 Then
        Err.Raise vbObjectError + 104, , "Nenhum objeto chamado " & NOME_BASE_ALINHAMENTO & " foi encontrado nos escaninhos importados."
    End If

    AlinharNoCentroTopoInterno estrutura, bases

End Sub

Private Sub AjustarIntermediariosCascata(ByVal estrutura As Shape, ByVal grupoEscaninhos As Shape, ByRef escPorColuna() As Long)

    Dim int1 As ShapeRange
    Dim linhaInferior As ShapeRange

    If MaiorValorArray(escPorColuna) > 10 Then Exit Sub

    RemoverObjetosPorNomeNoShape estrutura, NOME_INT2_CASC
    RemoverObjetosPorNomeNoShape estrutura, NOME_INT3_CASC

    Set int1 = New ShapeRange
    ColetarObjetosPorNomeNoShape estrutura, NOME_INT1_CASC, int1

    If int1.Count = 0 Then
        Err.Raise vbObjectError + 105, , "Nenhum objeto chamado " & NOME_INT1_CASC & " foi encontrado na estrutura."
    End If

    Set linhaInferior = ObterLinhaInferiorEscaninhos(grupoEscaninhos)

    If linhaInferior.Count = 0 Then
        Err.Raise vbObjectError + 106, , "Nao foi possivel localizar a linha inferior dos escaninhos."
    End If

    AlinharNoCentroFundoInterno int1, linhaInferior

End Sub

Private Sub AjustarLarguraObjetosInt(ByVal estrutura As Shape, ByVal grupoEscaninhos As Shape)

    Dim objetosInt As ShapeRange
    Dim intSupRange As ShapeRange
    Dim intDirRange As ShapeRange
    Dim intEsqRange As ShapeRange
    Dim shp As Shape
    Dim intSup As Shape

    Dim grupoX As Double
    Dim grupoY As Double
    Dim grupoW As Double
    Dim grupoH As Double
    Dim alvoEsquerda As Double
    Dim alvoLargura As Double

    grupoEscaninhos.GetBoundingBox grupoX, grupoY, grupoW, grupoH

    alvoEsquerda = grupoX - SOBRA_INT_LATERAL
    alvoLargura = grupoW + (SOBRA_INT_LATERAL * 2)

    Set objetosInt = New ShapeRange
    ColetarObjetosPorPrefixoNoShape estrutura, "INT", objetosInt

    For Each shp In objetosInt
        If UCase$(shp.Name) <> UCase$(NOME_INTDIR_ESTRUTURA) And _
           UCase$(shp.Name) <> UCase$(NOME_INTESQ_ESTRUTURA) Then

            RedimensionarLarguraMantendoAltura shp, alvoEsquerda, alvoLargura
        End If
    Next shp

    Set intSupRange = New ShapeRange
    ColetarObjetosPorNomeNoShape estrutura, NOME_INTSUP_ESTRUTURA, intSupRange

    If intSupRange.Count = 0 Then
        Err.Raise vbObjectError + 107, , "Nenhum objeto chamado " & NOME_INTSUP_ESTRUTURA & " foi encontrado na estrutura."
    End If

    Set intSup = intSupRange(1)

    Set intEsqRange = New ShapeRange
    ColetarObjetosPorNomeNoShape estrutura, NOME_INTESQ_ESTRUTURA, intEsqRange

    If intEsqRange.Count = 0 Then
        Err.Raise vbObjectError + 108, , "Nenhum objeto chamado " & NOME_INTESQ_ESTRUTURA & " foi encontrado na estrutura."
    End If

    Set intDirRange = New ShapeRange
    ColetarObjetosPorNomeNoShape estrutura, NOME_INTDIR_ESTRUTURA, intDirRange

    If intDirRange.Count = 0 Then
        Err.Raise vbObjectError + 109, , "Nenhum objeto chamado " & NOME_INTDIR_ESTRUTURA & " foi encontrado na estrutura."
    End If

    PosicionarObjetoAoLadoDoAlvo intEsqRange(1), intSup, False
    PosicionarObjetoAoLadoDoAlvo intDirRange(1), intSup, True

End Sub

Private Sub RedimensionarLarguraMantendoAltura(ByVal shp As Shape, ByVal alvoEsquerda As Double, ByVal alvoLargura As Double)

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    shp.GetBoundingBox x, y, w, h
    shp.SetSize alvoLargura, h

    shp.GetBoundingBox x, y, w, h
    shp.Move alvoEsquerda - x, 0

End Sub

Private Sub PosicionarObjetoAoLadoDoAlvo(ByVal shp As Shape, ByVal alvo As Shape, ByVal ladoDireito As Boolean)

    If ladoDireito Then
        shp.LeftX = alvo.RightX
    Else
        shp.RightX = alvo.LeftX
    End If

    AlinharTopoComoAtalhoT shp, alvo

End Sub

Private Sub AlinharTopoComoAtalhoT(ByVal shp As Shape, ByVal alvo As Shape)

    Dim sr As ShapeRange

    Set sr = New ShapeRange
    sr.Add shp
    sr.AlignToShape cdrAlignTop, alvo, cdrTextAlignBoundingBox

End Sub

Private Function ObterLinhaInferiorEscaninhos(ByVal grupoEscaninhos As Shape) As ShapeRange

    Dim escaninhos() As Shape
    Dim total As Long
    Dim i As Long

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double
    Dim menorY As Double
    Dim tolerancia As Double

    Set ObterLinhaInferiorEscaninhos = New ShapeRange

    ReDim escaninhos(1 To MAX_ESCANINHOS_DESENHO * LIMITE_COLUNAS_POR_QUADRO * 2)

    ColetarEscaninhos grupoEscaninhos, escaninhos, total

    If total = 0 Then Exit Function

    menorY = 999999
    tolerancia = 0.5

    For i = 1 To total
        escaninhos(i).GetBoundingBox x, y, w, h

        If y < menorY Then
            menorY = y
        End If
    Next i

    For i = 1 To total
        escaninhos(i).GetBoundingBox x, y, w, h

        If Abs(y - menorY) <= tolerancia Then
            ObterLinhaInferiorEscaninhos.Add escaninhos(i)
        End If
    Next i

End Function

Private Sub ColetarObjetosPorNomeNoShape(ByVal shp As Shape, ByVal nomeObjeto As String, ByVal resultado As ShapeRange)

    Dim filho As Shape

    If UCase$(shp.Name) = UCase$(nomeObjeto) Then
        resultado.Add shp
    End If

    If shp.Type = cdrGroupShape Then
        For Each filho In shp.Shapes
            ColetarObjetosPorNomeNoShape filho, nomeObjeto, resultado
        Next filho
    End If

End Sub

Private Sub ColetarObjetosPorPrefixoNoShape(ByVal shp As Shape, ByVal prefixo As String, ByVal resultado As ShapeRange)

    Dim filho As Shape

    If UCase$(Left$(shp.Name, Len(prefixo))) = UCase$(prefixo) Then
        resultado.Add shp
    End If

    If shp.Type = cdrGroupShape Then
        For Each filho In shp.Shapes
            ColetarObjetosPorPrefixoNoShape filho, prefixo, resultado
        Next filho
    End If

End Sub

Private Sub RemoverObjetosPorNomeNoShape(ByVal shp As Shape, ByVal nomeObjeto As String)

    Dim encontrados As ShapeRange
    Dim i As Long

    Set encontrados = New ShapeRange

    ColetarObjetosPorNomeNoShape shp, nomeObjeto, encontrados

    For i = encontrados.Count To 1 Step -1
        encontrados(i).Delete
    Next i

End Sub

Private Sub AlinharNoCentroTopoInterno(ByVal shp As Shape, ByVal alvo As ShapeRange)

    Dim alvoX As Double
    Dim alvoY As Double
    Dim alvoW As Double
    Dim alvoH As Double

    Dim shpX As Double
    Dim shpY As Double
    Dim shpW As Double
    Dim shpH As Double

    alvo.GetBoundingBox alvoX, alvoY, alvoW, alvoH
    shp.GetBoundingBox shpX, shpY, shpW, shpH

    shp.Move _
        (alvoX + (alvoW / 2)) - (shpX + (shpW / 2)), _
        (alvoY + alvoH) - (shpY + shpH)

End Sub

Private Sub AlinharNoCentroFundoInterno(ByVal shp As ShapeRange, ByVal alvo As ShapeRange)

    Dim alvoX As Double
    Dim alvoY As Double
    Dim alvoW As Double
    Dim alvoH As Double

    Dim shpX As Double
    Dim shpY As Double
    Dim shpW As Double
    Dim shpH As Double

    alvo.GetBoundingBox alvoX, alvoY, alvoW, alvoH
    shp.GetBoundingBox shpX, shpY, shpW, shpH

    shp.Move _
        (alvoX + (alvoW / 2)) - (shpX + (shpW / 2)), _
        alvoY - shpY

End Sub

Private Sub PosicionarPeloTopoEsquerdo(ByVal shp As Shape, ByVal alvoEsquerda As Double, ByVal alvoTopo As Double)

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    shp.GetBoundingBox x, y, w, h
    shp.Move alvoEsquerda - x, alvoTopo - (y + h)

End Sub

Private Function ObterDireita(ByVal shp As Shape) As Double

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    shp.GetBoundingBox x, y, w, h

    ObterDireita = x + w

End Function

Private Sub RemoverObjetosPorNome(ByVal nomeObjeto As String)

    Dim sr As ShapeRange
    Dim i As Long

    Set sr = ActivePage.FindShapes(Name:=nomeObjeto)

    If sr.Count = 0 Then Exit Sub

    For i = sr.Count To 1 Step -1
        sr(i).Delete
    Next i

End Sub

Private Function MaiorValorArray(ByRef valores() As Long) As Long

    Dim i As Long

    For i = LBound(valores) To UBound(valores)
        If valores(i) > MaiorValorArray Then
            MaiorValorArray = valores(i)
        End If
    Next i

End Function

Private Function EhInteiroPositivo(ByVal valor As String) As Boolean

    Dim i As Long
    Dim caractere As String

    valor = Trim$(valor)

    If Len(valor) = 0 Then Exit Function

    For i = 1 To Len(valor)
        caractere = Mid$(valor, i, 1)

        If caractere < "0" Or caractere > "9" Then
            EhInteiroPositivo = False
            Exit Function
        End If
    Next i

    EhInteiroPositivo = CLng(valor) > 0

End Function


