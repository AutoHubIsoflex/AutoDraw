Attribute VB_Name = "cavaleteAutoV2"
Option Explicit

Private Const TOLERANCIA_COR As Double = 0.5

Sub CavaleteCinza()
    InserirCavaleteMantendoLogica _
        "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\CAVALETES\CAVALETE_CZ.cdr", _
        "CAVALETE-METALON3-CZ"
End Sub

Sub CavaleteBranco()
    InserirCavaleteMantendoLogica _
        "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\CAVALETES\CAVALETE_BR.cdr", _
        "CAVALETE-METALON3-BR"
End Sub

Sub CavaletePreto()
    InserirCavaleteMantendoLogica _
        "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\CAVALETES\CAVALETE_PT.cdr", _
        "CAVALETE-METALON3-PT"
End Sub

Private Sub InserirCavaleteMantendoLogica(ByVal caminho As String, ByVal nomeGrupo As String)
    Dim quadro As Shape
    Dim cavalelete As Shape
    Dim offset As Double
    Dim offsetY As Double
    
    Dim s As Shape
    Dim grupo As Shape
    Dim filho As Shape
    
    Dim copiaGrupo As Shape
    
    Dim candidatos As Collection
    Dim maiorShape As Shape
    Dim maiorArea As Double
    Dim area As Double
    
    On Error GoTo TrataErro
    
    ' BUSCA AUTOMÁTICA DO RETÂNGULO MAGENTA
    Set candidatos = New Collection
    maiorArea = 0
    
    For Each s In ActivePage.Shapes
        If EhRetanguloMagenta(s) Then
            area = s.SizeWidth * s.SizeHeight
            If area > 1 Then
                candidatos.Add s
                If area > maiorArea Then
                    maiorArea = area
                    Set maiorShape = s
                End If
            End If
        End If
    Next s
    
    ' LÓGICA HÍBRIDA (SELEÇÃO MANUAL OU AUTOMÁTICA)
    If candidatos.Count = 0 Then
        MsgBox "Nenhum retângulo com borda magenta encontrado.", vbExclamation
        Exit Sub
    End If
    
    If candidatos.Count > 1 Then
        If ActiveSelection.Shapes.Count > 0 Then
            Set quadro = ActiveSelection.Shapes(1)
            
            If Not EhRetanguloMagenta(quadro) Then
                MsgBox "O objeto selecionado não é um retângulo com borda magenta.", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Mais de um retângulo com borda magenta encontrado. Selecione manualmente o quadro.", vbCritical
            Exit Sub
        End If
    Else
        Set quadro = maiorShape
    End If

    If Dir(caminho) = "" Then
        MsgBox "Arquivo não encontrado:" & vbCrLf & caminho, vbCritical
        Exit Sub
    End If

    ' Importa cavalete
    ActiveLayer.Import caminho
    
    If ActiveSelection Is Nothing Then
        MsgBox "Falha ao importar o cavalete: nenhuma seleção ativa foi criada.", vbCritical
        Exit Sub
    End If
    
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Falha ao importar o cavalete: nenhum objeto foi selecionado após importar.", vbCritical
        Exit Sub
    End If
    
    Set cavalelete = ActiveSelection.Shapes(1)

    ' Alinha topo
    cavalelete.TopY = quadro.TopY

    ' Alinha esquerda
    cavalelete.LeftX = quadro.LeftX

    ' Move 418,8 mm para a esquerda
    offset = ActiveDocument.ToUnits(418.8, cdrMillimeter)
    cavalelete.LeftX = cavalelete.LeftX - offset

    ' Sobe 30,4 mm
    offsetY = ActiveDocument.ToUnits(30.4, cdrMillimeter)
    cavalelete.TopY = cavalelete.TopY + offsetY

    ' Acha o grupo - MANTIDO COMO NO ORIGINAL
    For Each s In cavalelete.Shapes.All
        If s.Type = cdrGroupShape Then
            If s.Name = nomeGrupo Then
                Set grupo = s
                Exit For
            End If
        End If
    Next s

    If grupo Is Nothing Then
        MsgBox "Grupo '" & nomeGrupo & "' não encontrado."
        Exit Sub
    End If

    ' Acha maoFrancesa - agora com busca recursiva,
    ' mas sem alterar a lógica principal de posicionamento e duplicação
    Set filho = BuscarShapePorNomeRecursivo(grupo, "maoFrancesa")

    If filho Is Nothing Then
        MsgBox "Objeto 'maoFrancesa' não encontrado dentro do grupo '" & nomeGrupo & "'.", vbCritical
        Exit Sub
    End If

    ' Alinha na base do quadro
    filho.BottomY = quadro.BottomY

    ' Move 188,419 mm para baixo
    filho.BottomY = filho.BottomY - ActiveDocument.ToUnits(188.419, cdrMillimeter)

    ' ===== DUPLICAR E ESPELHAR =====
    ' MANTIDO EXATAMENTE COMO NA SUA LÓGICA ORIGINAL
    Set copiaGrupo = grupo.Duplicate

    ' Espelha horizontalmente
    copiaGrupo.Flip cdrFlipHorizontal

    ' Alinha na direita do quadro
    copiaGrupo.RightX = quadro.RightX

    ' Move 147 mm para a direita
    copiaGrupo.RightX = copiaGrupo.RightX + ActiveDocument.ToUnits(147, cdrMillimeter)

    Exit Sub

TrataErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub

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

Private Function BuscarShapePorNomeRecursivo(ByVal raiz As Shape, ByVal nomeBuscado As String) As Shape
    Dim s As Shape
    Dim encontrado As Shape
    
    On Error GoTo Falha
    
    If raiz Is Nothing Then Exit Function
    
    If StrComp(Trim$(raiz.Name), nomeBuscado, vbTextCompare) = 0 Then
        Set BuscarShapePorNomeRecursivo = raiz
        Exit Function
    End If
    
    If TemFilhos(raiz) Then
        For Each s In raiz.Shapes.All
            If StrComp(Trim$(s.Name), nomeBuscado, vbTextCompare) = 0 Then
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
    End If
    
    Exit Function

Falha:
    Set BuscarShapePorNomeRecursivo = Nothing
End Function

Private Function TemFilhos(ByVal s As Shape) As Boolean
    On Error GoTo Falha
    
    TemFilhos = False
    If s Is Nothing Then Exit Function
    
    TemFilhos = (s.Shapes.Count > 0)
    Exit Function

Falha:
    TemFilhos = False
End Function

