Attribute VB_Name = "cavaleteBeta1"
Sub CavaleteCinza()
    Dim quadro As Shape
    Dim cavalelete As Shape
    Dim caminho As String
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
    
    ' BUSCA AUTOMÁTICA DO RETÂNGULO MAGENTA
    
    Set candidatos = New Collection
    maiorArea = 0
    
    For Each s In ActivePage.Shapes
        If s.Type = cdrRectangleShape Then
            If Not s.Outline Is Nothing Then
                ' verifica borda magenta em CMYK (com tolerância)
                If Abs(s.Outline.Color.CMYKCyan - 0) < 0.5 And _
                   Abs(s.Outline.Color.CMYKMagenta - 100) < 0.5 And _
                   Abs(s.Outline.Color.CMYKYellow - 0) < 0.5 And _
                   Abs(s.Outline.Color.CMYKBlack - 0) < 0.5 Then
                   
                   area = s.SizeWidth * s.SizeHeight
                   If area > 1 Then
                       candidatos.Add s
                       If area > maiorArea Then
                           maiorArea = area
                           Set maiorShape = s
                       End If
                   End If
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
        ' mais de um magenta - usar seleção manual
        If ActiveSelection.Shapes.Count > 0 Then
            Set quadro = ActiveSelection.Shapes(1)
            
            ' valida se selecionado é magenta
            If Not (quadro.Type = cdrRectangleShape And _
                Not quadro.Outline Is Nothing And _
                Abs(quadro.Outline.Color.CMYKCyan - 0) < 0.5 And _
                Abs(quadro.Outline.Color.CMYKMagenta - 100) < 0.5 And _
                Abs(quadro.Outline.Color.CMYKYellow - 0) < 0.5 And _
                Abs(quadro.Outline.Color.CMYKBlack - 0) < 0.5) Then
                
                MsgBox "O objeto selecionado não é um retângulo com borda magenta.", vbExclamation
                Exit Sub
            End If
            
        Else
            MsgBox "Mais de um retângulo com borda magenta encontrado. Selecione manualmente o quadro.", vbCritical
            Exit Sub
        End If
        
    Else
        ' apenas um magenta - automático
        Set quadro = maiorShape
    End If

    caminho = "E:\ARQUIVOS DIVERSOS ISOFLEX (ESSENCIAIS)\NOVAS IDEIAS\AutoDraw\Simbolos\CAVALETES\CAVALETE_CZ.CDR"

    ' Importa cavalete
    ActiveLayer.Import caminho
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

    ' Acha o grupo
    For Each s In cavalelete.Shapes.All
        If s.Type = cdrGroupShape Then
            If s.Name = "CAVALETE-METALON3-CZ" Then
                Set grupo = s
                Exit For
            End If
        End If
    Next s

    If grupo Is Nothing Then
        MsgBox "Grupo não encontrado."
        Exit Sub
    End If

    ' Acha mao Francesa
    For Each filho In grupo.Shapes.All
        If filho.Name = "maoFrancesa" Then
            
            ' Alinha na base do quadro
            filho.BottomY = quadro.BottomY
            
            ' Move 188,419 mm para baixo
            filho.BottomY = filho.BottomY - ActiveDocument.ToUnits(188.419, cdrMillimeter)
            
            Exit For
        End If
    Next filho

    ' ===== DUPLICAR E ESPELHAR =====
    
    Set copiaGrupo = grupo.Duplicate

    ' Espelha horizontalmente
    copiaGrupo.Flip cdrFlipHorizontal

    ' Alinha na direita do quadro
    copiaGrupo.RightX = quadro.RightX

    ' Move 147 mm para a direita
    copiaGrupo.RightX = copiaGrupo.RightX + ActiveDocument.ToUnits(147, cdrMillimeter)

End Sub
Sub CavaleteBranco()
    Dim quadro As Shape
    Dim cavalelete As Shape
    Dim caminho As String
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
    
    ' BUSCA AUTOMÁTICA DO RETÂNGULO MAGENTA

    Set candidatos = New Collection
    maiorArea = 0
    
    For Each s In ActivePage.Shapes
        If s.Type = cdrRectangleShape Then
            If Not s.Outline Is Nothing Then
                ' verifica borda magenta em CMYK (com tolerância)
                If Abs(s.Outline.Color.CMYKCyan - 0) < 0.5 And _
                   Abs(s.Outline.Color.CMYKMagenta - 100) < 0.5 And _
                   Abs(s.Outline.Color.CMYKYellow - 0) < 0.5 And _
                   Abs(s.Outline.Color.CMYKBlack - 0) < 0.5 Then
                   
                   area = s.SizeWidth * s.SizeHeight
                   If area > 1 Then
                       candidatos.Add s
                       If area > maiorArea Then
                           maiorArea = area
                           Set maiorShape = s
                       End If
                   End If
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
        ' mais de um magenta - usar seleção manual
        If ActiveSelection.Shapes.Count > 0 Then
            Set quadro = ActiveSelection.Shapes(1)
            
            ' valida se selecionado é magenta
            If Not (quadro.Type = cdrRectangleShape And _
                Not quadro.Outline Is Nothing And _
                Abs(quadro.Outline.Color.CMYKCyan - 0) < 0.5 And _
                Abs(quadro.Outline.Color.CMYKMagenta - 100) < 0.5 And _
                Abs(quadro.Outline.Color.CMYKYellow - 0) < 0.5 And _
                Abs(quadro.Outline.Color.CMYKBlack - 0) < 0.5) Then
                
                MsgBox "O objeto selecionado não é um retângulo com borda magenta.", vbExclamation
                Exit Sub
            End If
            
        Else
            MsgBox "Mais de um retângulo com borda magenta encontrado. Selecione manualmente o quadro.", vbCritical
            Exit Sub
        End If
        
    Else
        ' apenas um magenta - automático
        Set quadro = maiorShape
    End If

    caminho = "E:\ARQUIVOS DIVERSOS ISOFLEX (ESSENCIAIS)\NOVAS IDEIAS\AutoDraw\Simbolos\CAVALETES\CAVALETE_BR.CDR"

    ' Importa cavalete
    ActiveLayer.Import caminho
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

    ' Acha o grupo
    For Each s In cavalelete.Shapes.All
        If s.Type = cdrGroupShape Then
            If s.Name = "CAVALETE-METALON3-BR" Then
                Set grupo = s
                Exit For
            End If
        End If
    Next s

    If grupo Is Nothing Then
        MsgBox "Grupo não encontrado."
        Exit Sub
    End If

    ' Acha mao Francesa
    For Each filho In grupo.Shapes.All
        If filho.Name = "maoFrancesa" Then
            
            ' Alinha na base do quadro
            filho.BottomY = quadro.BottomY
            
            ' Move 188,419 mm para baixo
            filho.BottomY = filho.BottomY - ActiveDocument.ToUnits(188.419, cdrMillimeter)
            
            Exit For
        End If
    Next filho

    ' ===== DUPLICAR E ESPELHAR =====
    
    Set copiaGrupo = grupo.Duplicate

    ' Espelha horizontalmente
    copiaGrupo.Flip cdrFlipHorizontal

    ' Alinha na direita do quadro
    copiaGrupo.RightX = quadro.RightX

    ' Move 147 mm para a direita
    copiaGrupo.RightX = copiaGrupo.RightX + ActiveDocument.ToUnits(147, cdrMillimeter)

End Sub
Sub CavaletePreto()
    Dim quadro As Shape
    Dim cavalelete As Shape
    Dim caminho As String
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
    
    ' BUSCA AUTOMÁTICA DO RETÂNGULO MAGENTA

    Set candidatos = New Collection
    maiorArea = 0
    
    For Each s In ActivePage.Shapes
        If s.Type = cdrRectangleShape Then
            If Not s.Outline Is Nothing Then
                ' verifica borda magenta em CMYK (com tolerância)
                If Abs(s.Outline.Color.CMYKCyan - 0) < 0.5 And _
                   Abs(s.Outline.Color.CMYKMagenta - 100) < 0.5 And _
                   Abs(s.Outline.Color.CMYKYellow - 0) < 0.5 And _
                   Abs(s.Outline.Color.CMYKBlack - 0) < 0.5 Then
                   
                   area = s.SizeWidth * s.SizeHeight
                   If area > 1 Then
                       candidatos.Add s
                       If area > maiorArea Then
                           maiorArea = area
                           Set maiorShape = s
                       End If
                   End If
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
        ' mais de um magenta - usar seleção manual
        If ActiveSelection.Shapes.Count > 0 Then
            Set quadro = ActiveSelection.Shapes(1)
            
            ' valida se selecionado é magenta
            If Not (quadro.Type = cdrRectangleShape And _
                Not quadro.Outline Is Nothing And _
                Abs(quadro.Outline.Color.CMYKCyan - 0) < 0.5 And _
                Abs(quadro.Outline.Color.CMYKMagenta - 100) < 0.5 And _
                Abs(quadro.Outline.Color.CMYKYellow - 0) < 0.5 And _
                Abs(quadro.Outline.Color.CMYKBlack - 0) < 0.5) Then
                
                MsgBox "O objeto selecionado não é um retângulo com borda magenta.", vbExclamation
                Exit Sub
            End If
            
        Else
            MsgBox "Mais de um retângulo com borda magenta encontrado. Selecione manualmente o quadro.", vbCritical
            Exit Sub
        End If
        
    Else
        ' apenas um magenta - automático
        Set quadro = maiorShape
    End If

    caminho = "E:\ARQUIVOS DIVERSOS ISOFLEX (ESSENCIAIS)\NOVAS IDEIAS\AutoDraw\Simbolos\CAVALETES\CAVALETE_PT.CDR"

    ' Importa cavalete
    ActiveLayer.Import caminho
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

    ' Acha o grupo
    For Each s In cavalelete.Shapes.All
        If s.Type = cdrGroupShape Then
            If s.Name = "CAVALETE-METALON3-PT" Then
                Set grupo = s
                Exit For
            End If
        End If
    Next s

    If grupo Is Nothing Then
        MsgBox "Grupo não encontrado."
        Exit Sub
    End If

    ' Acha mao Francesa
    For Each filho In grupo.Shapes.All
        If filho.Name = "maoFrancesa" Then
            
            ' Alinha na base do quadro
            filho.BottomY = quadro.BottomY
            
            ' Move 188,419 mm para baixo
            filho.BottomY = filho.BottomY - ActiveDocument.ToUnits(188.419, cdrMillimeter)
            
            Exit For
        End If
    Next filho

    ' ===== DUPLICAR E ESPELHAR =====
    
    Set copiaGrupo = grupo.Duplicate

    ' Espelha horizontalmente
    copiaGrupo.Flip cdrFlipHorizontal

    ' Alinha na direita do quadro
    copiaGrupo.RightX = quadro.RightX

    ' Move 147 mm para a direita
    copiaGrupo.RightX = copiaGrupo.RightX + ActiveDocument.ToUnits(147, cdrMillimeter)

End Sub


