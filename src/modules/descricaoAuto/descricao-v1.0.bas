Attribute VB_Name = "descricao"
Sub descricaoAuto()

    Dim sr As ShapeRange
    Dim sh As Shape
    Dim rect As Shape
    Dim txt As Shape
    
    Dim largura As Double
    Dim altura As Double
    Dim textoFinal As String
    
    ' =========================
    ' Variáveis de acessórios
    ' =========================
    Dim qtdKSIR_AD As Long
    Dim qtdKSIP_AD As Long
    Dim qtdKSIR_MG As Long
    Dim qtdKSIP_MG As Long
    Dim textoAcessorios As String
    
    ' =========================
    ' Controle de quadro magnético
    ' =========================
    Dim ehMagnetico As Boolean
    ehMagnetico = False
    
    ' Verifica se há seleção
    If ActiveSelectionRange.Count < 2 Then
        MsgBox "Selecione o retângulo e o texto.", vbExclamation
        Exit Sub
    End If
    
    Set sr = ActiveSelectionRange
    
    ' Identifica o retângulo e o texto
    For Each sh In sr
        If sh.Type = cdrRectangleShape Then
            Set rect = sh
        ElseIf sh.Type = cdrTextShape Then
            Set txt = sh
        End If
    Next sh
    
    ' Verifica se ambos foram encontrados
    If rect Is Nothing Or txt Is Nothing Then
        MsgBox "É necessário selecionar um retângulo e um texto.", vbCritical
        Exit Sub
    End If
    
    ' Garante unidade em milímetros
    ActiveDocument.Unit = cdrMillimeter
    
    ' Captura largura e altura
    largura = Round(rect.SizeWidth, 0)
    altura = Round(rect.SizeHeight, 0)
    
    ' =========================
    ' Leitura dos acessórios na página
    ' =========================
    qtdKSIR_AD = 0
    qtdKSIP_AD = 0
    qtdKSIR_MG = 0
    qtdKSIP_MG = 0
    
    For Each sh In ActivePage.Shapes
        
        ' Se existir MG no nome ? quadro magnético
        If InStr(1, sh.Name, "-MG", vbTextCompare) > 0 Then
            ehMagnetico = True
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
        End Select
    Next sh
    
    ' =========================
    ' Monta texto dos acessórios
    ' =========================
    textoAcessorios = ""
    
    If qtdKSIR_AD > 0 Then textoAcessorios = textoAcessorios & "- " & qtdKSIR_AD & " KSIR-A4-AD" & vbCrLf
    If qtdKSIP_AD > 0 Then textoAcessorios = textoAcessorios & "- " & qtdKSIP_AD & " KSIP-A4-AD" & vbCrLf
    If qtdKSIR_MG > 0 Then textoAcessorios = textoAcessorios & "- " & qtdKSIR_MG & " KSIR-A4-MG" & vbCrLf
    If qtdKSIP_MG > 0 Then textoAcessorios = textoAcessorios & "- " & qtdKSIP_MG & " KSIP-A4-MG" & vbCrLf
    
    ' =========================
    ' Montagem do texto final
    ' =========================
    
    If ehMagnetico = True Then
        ' TEXTO PARA QUADRO MAGNÉTICO
        textoFinal = "QUADRO BRANCO MAGNÉTICO" & vbCrLf & _
                     "PARA ESCRITA COM IMPRESSÃO " & vbCrLf & _
                     "DIGITAL UV. E LAMINAÇÃO PYT" & vbCrLf & _
                     "MED " & altura & "x" & largura & " - QPMM"
    Else
        ' TEXTO PARA QUADRO NORMAL
        textoFinal = "QUADRO BRANCO PARA ESCRITA" & vbCrLf & _
                     "COM IMPRESSÃO DIGITAL UV. E" & vbCrLf & _
                     "LAMINAÇÃO PYT MED " & altura & "x" & largura & vbCrLf & _
                     "- QPMS"
    End If
    
    ' Só adiciona acessórios se existir algum
    If textoAcessorios <> "" Then
        textoFinal = textoFinal & vbCrLf & vbCrLf & _
                     "ACESSÓRIOS:" & vbCrLf & vbCrLf & _
                     textoAcessorios
    End If
    
    ' Atualiza o texto
    txt.Text.Story = textoFinal
    
    MsgBox "Texto atualizado com sucesso!", vbInformation

End Sub


