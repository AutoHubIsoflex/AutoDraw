Attribute VB_Name = "Core"
Option Explicit

' =========================================================
' CORE - FLUXO DE INSERÇÃO DE CAVALETE
' =========================================================

Public Sub InserirCavalete(ByVal caminhoArquivo As String, ByVal nomeGrupo As String)
    Dim quadro As Shape
    Dim shapeImportado As Shape
    Dim grupoCavalete As Shape
    Dim maoFrancesa As Shape
    Dim grupoEspelhado As Shape

    On Error GoTo TrataErro

    Set quadro = ObterQuadroMagentaValido()
    If quadro Is Nothing Then Exit Sub

    If DeveAlertarMaoFrancesaInvertida(quadro) Then
        MsgBox "Cavalete com mão francesa invertida", vbExclamation
    ElseIf DeveAlertarCavaleteEspecial(quadro) Then
        MsgBox "A altura do quadro exige que o cavalete tenha uma medida especial. Ajuste manualmente.", vbExclamation
    End If

    If Not ArquivoExiste(caminhoArquivo) Then
        MsgBox "Arquivo não encontrado:" & vbCrLf & caminhoArquivo, vbCritical
        Exit Sub
    End If

    Set shapeImportado = ImportarArquivoCavalete(caminhoArquivo)
    If shapeImportado Is Nothing Then Exit Sub

    PosicionarCavaleteInicial shapeImportado, quadro

    Set grupoCavalete = ObterGrupoPorNome(shapeImportado, nomeGrupo)
    If grupoCavalete Is Nothing Then
        MsgBox "Grupo '" & nomeGrupo & "' não encontrado.", vbCritical
        Exit Sub
    End If

    Set maoFrancesa = BuscarShapePorNomeRecursivo(grupoCavalete, NOME_SHAPE_MAO_FRANCESA)
    If DeveExcluirMaoFrancesa(quadro) Then
        If Not maoFrancesa Is Nothing Then
            maoFrancesa.Delete
            Set maoFrancesa = Nothing
        End If
    Else
        If maoFrancesa Is Nothing Then
            MsgBox "Objeto '" & NOME_SHAPE_MAO_FRANCESA & "' não encontrado dentro do grupo '" & nomeGrupo & "'.", vbCritical
            Exit Sub
        End If

        PosicionarMaoFrancesa maoFrancesa, quadro
    End If

    Set grupoEspelhado = grupoCavalete.Duplicate
    EspelharEPosicionarGrupo grupoEspelhado, quadro

    Exit Sub

TrataErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Function DeveExcluirMaoFrancesa(ByVal quadro As Shape) As Boolean
    If quadro Is Nothing Then Exit Function

    DeveExcluirMaoFrancesa = (quadro.SizeHeight >= MmParaDocumento(ALTURA_MINIMA_EXCLUIR_MAO_FRANCESA_MM))
End Function

Private Function DeveAlertarMaoFrancesaInvertida(ByVal quadro As Shape) As Boolean
    Dim alturaQuadro As Double
    Dim limiteInferior As Double
    Dim limiteSuperior As Double

    If quadro Is Nothing Then Exit Function

    alturaQuadro = quadro.SizeHeight
    limiteInferior = MmParaDocumento(ALTURA_MINIMA_EXCLUIR_MAO_FRANCESA_MM)
    limiteSuperior = MmParaDocumento(ALTURA_ALERTA_CAVALETE_ESPECIAL_MM)

    DeveAlertarMaoFrancesaInvertida = (alturaQuadro >= limiteInferior And alturaQuadro <= limiteSuperior)
End Function

Private Function DeveAlertarCavaleteEspecial(ByVal quadro As Shape) As Boolean
    If quadro Is Nothing Then Exit Function

    DeveAlertarCavaleteEspecial = (quadro.SizeHeight > MmParaDocumento(ALTURA_ALERTA_CAVALETE_ESPECIAL_MM))
End Function

