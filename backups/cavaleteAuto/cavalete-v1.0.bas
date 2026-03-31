Attribute VB_Name = "cavaleteBeta1"
Sub InserirCavaleleteAlinharTopo()

    Dim quadro As Shape
    Dim cavalelete As Shape
    Dim caminho As String
    Dim offset As Double
    Dim offsetY As Double
    
    Dim s As Shape
    Dim grupo As Shape
    Dim filho As Shape
    
    Dim copiaGrupo As Shape

    caminho = "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\CAVALETES\CAVALETE_CZ.CDR"

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Selecione o retângulo do quadro primeiro."
        Exit Sub
    End If

    Set quadro = ActiveSelection.Shapes(1)

    ' Importa cavalelete
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

    ' ===== PARTE DO AJUSTE =====
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

    ' Acha "maoFrancesa"
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

    ' Move 418,8 mm para a direita
    copiaGrupo.RightX = copiaGrupo.RightX + ActiveDocument.ToUnits(418.8, cdrMillimeter)

End Sub
