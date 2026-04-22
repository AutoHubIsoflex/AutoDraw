Attribute VB_Name = "modMain"
' modMain
Option Explicit

Sub descricaoAuto()

    Dim tipoQuadro As tipoQuadro
    tipoQuadro = modUi.SolicitarTipoQuadro()
    If tipoQuadro = -1 Then Exit Sub

    Dim rect As Shape
    If Not modLayout.ObterRetanguloMagenta(rect) Then Exit Sub

    Dim txt As Shape
    If Not modLayout.TentarObterTextoSelecionado(txt) Then
        MsgBox "Selecione o texto de descrição e rode novamente.", vbExclamation
        Exit Sub
    End If

    ActiveDocument.Unit = cdrMillimeter
    Dim largura As Double: largura = Round(rect.SizeWidth, 0)
    Dim altura  As Double: altura = Round(rect.SizeHeight, 0)

    Dim catalogo As Collection
    Set catalogo = modCatalogo.CriarCatalogoAcessorios()

    Dim indice As Object
    Set indice = modCatalogo.CriarIndiceAcessorios(catalogo)

    Dim ehMG As Boolean
    Dim ehAD As Boolean
    Dim contadores As Object
    Dim medidasAcessorios As Object
    Set contadores = modLayout.ColetarAcessorios(indice, tipoQuadro, ehMG, ehAD, medidasAcessorios)

    If Not modUi.ConfirmarCompatibilidade(tipoQuadro, ehMG, ehAD) Then Exit Sub

    txt.Text.Story = modDescricao.MontarTextoCompleto(tipoQuadro, altura, largura, catalogo, contadores, medidasAcessorios)
    MsgBox "Texto atualizado com sucesso!", vbInformation

End Sub
