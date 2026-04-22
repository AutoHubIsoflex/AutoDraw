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
    If maoFrancesa Is Nothing Then
        MsgBox "Objeto '" & NOME_SHAPE_MAO_FRANCESA & "' não encontrado dentro do grupo '" & nomeGrupo & "'.", vbCritical
        Exit Sub
    End If

    PosicionarMaoFrancesa maoFrancesa, quadro

    Set grupoEspelhado = grupoCavalete.Duplicate
    EspelharEPosicionarGrupo grupoEspelhado, quadro

    Exit Sub

TrataErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub
