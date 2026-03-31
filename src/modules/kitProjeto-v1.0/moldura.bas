Attribute VB_Name = "molduraAutoV211"
Option Explicit

Sub molduraAzul()
    AplicarMoldura "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\MOLDURA AUTO\molduraAzul.cdr"
End Sub

Sub molduraCinza()
    AplicarMoldura "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\MOLDURA AUTO\molduraCinza.cdr"
End Sub

Sub molduraPreto()
    AplicarMoldura "E:\Desenvolvimento - Projeto\AutoHub\AutoDraw\assets\symbols\MOLDURA AUTO\molduraPreto.cdr"
End Sub

Private Sub AplicarMoldura(ByVal caminho As String)

    On Error GoTo TrataErro

    Dim s As Shape
    Dim candidatoSelecionado As Shape
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
    Dim candidatos As Collection
    Dim maiorShape As Shape
    Dim maiorArea As Double
    Dim area As Double

    offset = ActiveDocument.ToUnits(5.46, cdrMillimeter)

    Set candidatos = New Collection
    maiorArea = 0

    ' ===== BUSCA AUTOMÁTICA DO RETÂNGULO =====
    For Each s In ActivePage.Shapes

        If s.Type = cdrRectangleShape Then
            If ShapeTemContornoMagenta(s) Then

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

    Next s

    If candidatos.Count = 0 Then
        MsgBox "Nenhum retângulo com borda magenta encontrado.", vbExclamation
        Exit Sub
    End If

    ' ===== LÓGICA HÍBRIDA =====
    If candidatos.Count > 1 Then

        If ActiveSelectionRange.Count > 0 Then
            Set candidatoSelecionado = ActiveSelectionRange(1)

            If candidatoSelecionado.Type <> cdrRectangleShape Then
                MsgBox "O objeto selecionado não é um retângulo.", vbExclamation
                Exit Sub
            End If

            If Not ShapeTemContornoMagenta(candidatoSelecionado) Then
                MsgBox "O objeto selecionado não possui borda magenta CMYK válida.", vbExclamation
                Exit Sub
            End If

            Set s = candidatoSelecionado
        Else
            MsgBox "Mais de um retângulo magenta encontrado. Selecione manualmente o retângulo desejado.", vbCritical
            Exit Sub
        End If

    Else
        Set s = maiorShape
    End If

    ' ===== ARQUIVO =====
    If Dir(caminho) = "" Then
        MsgBox "Arquivo não encontrado: " & vbCrLf & caminho, vbCritical
        Exit Sub
    End If

    ActiveLayer.Import caminho

    If ActiveSelectionRange.Count = 0 Then
        MsgBox "O arquivo foi importado, mas nenhum objeto ficou selecionado após a importação.", vbCritical
        Exit Sub
    End If

    Set grupo = ActiveSelectionRange(1)

    If grupo Is Nothing Then
        MsgBox "Não foi possível obter o grupo importado.", vbCritical
        Exit Sub
    End If

    If grupo.Shapes.Count = 0 Then
        MsgBox "O objeto importado não contém shapes internos.", vbCritical
        Exit Sub
    End If

    ' ===== LOCALIZA PEÇAS =====
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
        MsgBox "Alguma cantoneira não foi encontrada no arquivo importado.", vbCritical
        Exit Sub
    End If

    If tuboDir Is Nothing Or tuboSup Is Nothing Or tuboEsq Is Nothing Or tuboInf Is Nothing Then
        MsgBox "Algum tubo não foi encontrado no arquivo importado.", vbCritical
        Exit Sub
    End If

    ' ===== POSICIONAMENTO DAS CANTONEIRAS =====

    ' SUPERIOR DIREITA
    cantSupDir.LeftX = s.RightX - cantSupDir.SizeWidth + offset
    cantSupDir.TopY = s.TopY + offset

    ' SUPERIOR ESQUERDA
    cantSupEsq.LeftX = s.LeftX - offset
    cantSupEsq.TopY = s.TopY + offset

    ' INFERIOR ESQUERDA
    cantInfEsq.LeftX = s.LeftX - offset
    cantInfEsq.TopY = s.BottomY + cantInfEsq.SizeHeight - offset

    ' INFERIOR DIREITA
    cantInfDir.LeftX = s.RightX - cantInfDir.SizeWidth + offset
    cantInfDir.TopY = s.BottomY + cantInfDir.SizeHeight - offset

    ' ===== POSICIONAMENTO INICIAL DOS TUBOS =====

    ' TUBO DIREITO
    tuboDir.LeftX = cantSupDir.RightX - tuboDir.SizeWidth
    tuboDir.CenterY = s.CenterY

    ' TUBO SUPERIOR
    tuboSup.CenterX = s.CenterX
    tuboSup.TopY = cantSupDir.TopY

    ' TUBO ESQUERDO
    tuboEsq.LeftX = cantSupEsq.LeftX
    tuboEsq.CenterY = s.CenterY

    ' TUBO INFERIOR
    tuboInf.CenterX = s.CenterX
    tuboInf.BottomY = cantInfDir.BottomY

    ' ===== AJUSTE TUBO INFERIOR =====
    Dim xEsq As Double
    Dim xDir As Double
    Dim novaLargura As Double

    xEsq = cantInfEsq.CenterX
    xDir = cantInfDir.CenterX
    novaLargura = xDir - xEsq

    If novaLargura > 0 Then
        tuboInf.SetSize novaLargura, tuboInf.SizeHeight
        tuboInf.CenterX = (xEsq + xDir) / 2
    End If

    ' ===== AJUSTE TUBO ESQUERDO =====
    Dim ySupEsq As Double
    Dim yInfEsq As Double
    Dim novaAlturaEsq As Double

    ySupEsq = cantSupEsq.CenterY
    yInfEsq = cantInfEsq.CenterY
    novaAlturaEsq = ySupEsq - yInfEsq

    If novaAlturaEsq > 0 Then
        tuboEsq.SetSize tuboEsq.SizeWidth, novaAlturaEsq
        tuboEsq.CenterY = (ySupEsq + yInfEsq) / 2
    End If

    ' ===== AJUSTE TUBO DIREITO =====
    Dim ySupDir As Double
    Dim yInfDir As Double
    Dim novaAlturaDir As Double

    ySupDir = cantSupDir.CenterY
    yInfDir = cantInfDir.CenterY
    novaAlturaDir = ySupDir - yInfDir

    If novaAlturaDir > 0 Then
        tuboDir.SetSize tuboDir.SizeWidth, novaAlturaDir
        tuboDir.CenterY = (ySupDir + yInfDir) / 2
    End If

    ' ===== AJUSTE TUBO SUPERIOR =====
    Dim xEsqSup As Double
    Dim xDirSup As Double
    Dim novaLarguraSup As Double

    xEsqSup = cantSupEsq.CenterX
    xDirSup = cantSupDir.CenterX
    novaLarguraSup = xDirSup - xEsqSup

    If novaLarguraSup > 0 Then
        tuboSup.SetSize novaLarguraSup, tuboSup.SizeHeight
        tuboSup.CenterX = (xEsqSup + xDirSup) / 2
    End If

    Exit Sub

TrataErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "AplicarMoldura"
End Sub

Private Function ShapeTemContornoMagenta(ByVal s As Shape) As Boolean
    On Error GoTo Falha

    ShapeTemContornoMagenta = False

    If s Is Nothing Then Exit Function
    If s.Outline Is Nothing Then Exit Function

    ShapeTemContornoMagenta = CorEhMagentaCMYK(s.Outline.Color)
    Exit Function

Falha:
    ShapeTemContornoMagenta = False
    Err.Clear
End Function

Private Function CorEhMagentaCMYK(ByVal cor As Color) As Boolean
    On Error GoTo Falha

    CorEhMagentaCMYK = _
        Abs(cor.CMYKCyan - 0) < 0.5 And _
        Abs(cor.CMYKMagenta - 100) < 0.5 And _
        Abs(cor.CMYKYellow - 0) < 0.5 And _
        Abs(cor.CMYKBlack - 0) < 0.5

    Exit Function

Falha:
    ' Se a cor não for CMYK compatível, simplesmente retorna False
    CorEhMagentaCMYK = False
    Err.Clear
End Function

