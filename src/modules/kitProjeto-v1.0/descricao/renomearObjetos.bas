Attribute VB_Name = "renomearObjetos"
Sub RenomearObjetosSelecionados_Igual()

    Dim sr As ShapeRange
    Dim nomeBase As String
    Dim i As Integer

    ' Verifica se há seleção
    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Selecione pelo menos um objeto.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' Solicita o nome
    nomeBase = InputBox("Digite o nome para os objetos:", "Renomear Objetos")

    ' Validação
    If Trim(nomeBase) = "" Then
        MsgBox "Nenhum nome foi informado.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' Pega seleção
    Set sr = ActiveSelectionRange

    ' Renomeia todos com o mesmo nome
    For i = 1 To sr.Count
        sr(i).Name = nomeBase
    Next i

    MsgBox sr.Count & " objeto(s) renomeado(s) com sucesso.", vbInformation, "Concluído"

End Sub
