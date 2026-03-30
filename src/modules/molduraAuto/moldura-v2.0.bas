Attribute VB_Name = "molduraBeta1"
Sub molduraAuto()

    Dim sr As ShapeRange
    Dim s As Shape
    Dim caminho As String
    
    Dim grupo As Shape
    Dim sh As Shape
    Dim cantSupDir As Shape
    Dim cantSupEsq As Shape
    Dim cantInfEsq As Shape
    Dim cantInfDir As Shape
    Dim tuboDir As Shape
    Dim tuboSup As Shape
    Dim tuboEsq As Shape
    Dim tuboInf As Shape
    Dim offset As Double

    ' OFFSET EM MILÍMETROS REAL
    offset = ActiveDocument.ToUnits(5.46, cdrMillimeter)

    ' ================================
    ' BUSCA AUTOMÁTICA DO RETÂNGULO
    ' ================================
    
    Dim candidatos As Collection
    Set candidatos = New Collection

    Dim maiorShape As Shape
    Dim maiorArea As Double
    maiorArea = 0

    Dim area As Double

    ' percorre todos os shapes da página
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

    ' nenhum encontrado
    If candidatos.Count = 0 Then
        MsgBox "Nenhum retângulo com borda magenta encontrado.", vbExclamation
        Exit Sub
    End If

    ' ================================
    ' LÓGICA HÍBRIDA (AUTO + MANUAL)
    ' ================================

    If candidatos.Count > 1 Then
        
        ' se usuário selecionou algo ? usa isso
        If ActiveSelection.Shapes.Count > 0 Then
            
            Set s = ActiveSelection.Shapes(1)
            
            ' valida se é magenta
            If Not (Abs(s.Outline.Color.CMYKCyan - 0) < 0.5 And _
                    Abs(s.Outline.Color.CMYKMagenta - 100) < 0.5 And _
                    Abs(s.Outline.Color.CMYKYellow - 0) < 0.5 And _
                    Abs(s.Outline.Color.CMYKBlack - 0) < 0.5) Then
                
                MsgBox "O objeto selecionado não é magenta.", vbExclamation
                Exit Sub
            End If
            
        Else
            MsgBox "Mais de um magenta encontrado. Selecione manualmente o magenta.", vbCritical
            Exit Sub
        End If
        
    Else
        ' apenas um ? automático
        Set s = maiorShape
    End If

    ' ================================
    ' RESTANTE DO SEU CÓDIGO (INALTERADO)
    ' ================================

    caminho = "E:\ARQUIVOS DIVERSOS ISOFLEX (ESSENCIAIS)\NOVAS IDEIAS\AutoDraw\Simbolos\MOLDURA AUTO\1210X1830mm.cdr"

    If dir(caminho) = "" Then
        MsgBox "Arquivo não encontrado!"
        Exit Sub
    End If

    ActiveLayer.Import caminho

    Set grupo = ActiveSelectionRange(1)

    For Each sh In grupo.Shapes
        Select Case sh.Name
            Case "cantSupDir": Set cantSupDir = sh
            Case "cantSupEsq": Set cantSupEsq = sh
            Case "cantInfEsq": Set cantInfEsq = sh
            Case "cantInfDir": Set cantInfDir = sh
            Case "tuboDir": Set tuboDir = sh
            Case "tuboSup": Set tuboSup = sh
            Case "tuboEsq": Set tuboEsq = sh
            Case "tuboInf": Set tuboInf = sh
        End Select
    Next sh

    If cantSupDir Is Nothing Or cantSupEsq Is Nothing Or cantInfEsq Is Nothing Or cantInfDir Is Nothing Then
        MsgBox "Alguma cantoneira não foi encontrada!"
        Exit Sub
    End If

    If tuboDir Is Nothing Or tuboSup Is Nothing Or tuboEsq Is Nothing Or tuboInf Is Nothing Then
        MsgBox "Algum tubo não foi encontrado!"
        Exit Sub
    End If

    ' ===== SUPERIOR DIREITA =====
    cantSupDir.LeftX = s.RightX - cantSupDir.SizeWidth
    cantSupDir.TopY = s.TopY
    cantSupDir.LeftX = cantSupDir.LeftX + offset
    cantSupDir.TopY = cantSupDir.TopY + offset

    ' ===== SUPERIOR ESQUERDA =====
    cantSupEsq.LeftX = s.LeftX
    cantSupEsq.TopY = s.TopY
    cantSupEsq.LeftX = cantSupEsq.LeftX - offset
    cantSupEsq.TopY = cantSupEsq.TopY + offset

    ' ===== INFERIOR ESQUERDA =====
    cantInfEsq.LeftX = s.LeftX
    cantInfEsq.TopY = s.BottomY + cantInfEsq.SizeHeight
    cantInfEsq.LeftX = cantInfEsq.LeftX - offset
    cantInfEsq.TopY = cantInfEsq.TopY - offset

    ' ===== INFERIOR DIREITA =====
    cantInfDir.LeftX = s.RightX - cantInfDir.SizeWidth
    cantInfDir.TopY = s.BottomY + cantInfDir.SizeHeight
    cantInfDir.LeftX = cantInfDir.LeftX + offset
    cantInfDir.TopY = cantInfDir.TopY - offset

    ' ===== TUBO DIREITO =====
    tuboDir.LeftX = cantSupDir.RightX - tuboDir.SizeWidth
    tuboDir.CenterY = s.CenterY

    ' ===== TUBO SUPERIOR =====
    tuboSup.CenterX = s.CenterX
    tuboSup.TopY = cantSupDir.TopY

    ' ===== TUBO ESQUERDO =====
    tuboEsq.LeftX = cantSupEsq.LeftX
    tuboEsq.CenterY = s.CenterY

    ' ===== TUBO INFERIOR =====
    tuboInf.CenterX = s.CenterX
    tuboInf.BottomY = cantInfDir.BottomY

    ' ===== AJUSTE DE LARGURA DO TUBO INFERIOR =====

    Dim xEsq As Double
    Dim xDir As Double
    Dim novaLargura As Double
    
    xEsq = cantInfEsq.CenterX
    xDir = cantInfDir.CenterX
    
    novaLargura = xDir - xEsq
    
    tuboInf.SetSize novaLargura, tuboInf.SizeHeight
    tuboInf.CenterX = (xEsq + xDir) / 2

    ' ===== AJUSTE TUBO ESQUERDO =====

    Dim ySupEsq As Double
    Dim yInfEsq As Double
    Dim novaAlturaEsq As Double
    
    ySupEsq = cantSupEsq.CenterY
    yInfEsq = cantInfEsq.CenterY
    
    novaAlturaEsq = ySupEsq - yInfEsq
    
    tuboEsq.SetSize tuboEsq.SizeWidth, novaAlturaEsq
    tuboEsq.CenterY = (ySupEsq + yInfEsq) / 2
    
    ' ===== AJUSTE TUBO DIREITO =====
    
    Dim ySupDir As Double
    Dim yInfDir As Double
    Dim novaAlturaDir As Double
    
    ySupDir = cantSupDir.CenterY
    yInfDir = cantInfDir.CenterY
    
    novaAlturaDir = ySupDir - yInfDir
    
    tuboDir.SetSize tuboDir.SizeWidth, novaAlturaDir
    tuboDir.CenterY = (ySupDir + yInfDir) / 2
    
    ' ===== AJUSTE TUBO SUPERIOR =====
    
    Dim xEsqSup As Double
    Dim xDirSup As Double
    Dim novaLarguraSup As Double
    
    xEsqSup = cantSupEsq.CenterX
    xDirSup = cantSupDir.CenterX
    
    novaLarguraSup = xDirSup - xEsqSup
    
    tuboSup.SetSize novaLarguraSup, tuboSup.SizeHeight
    tuboSup.CenterX = (xEsqSup + xDirSup) / 2

End Sub


