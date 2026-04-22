Attribute VB_Name = "Map"
Option Explicit

' =========================================================
' BUSCA DE SHAPES E GRUPOS
' =========================================================

Public Function ObterGrupoPorNome(ByVal shapeRaiz As Shape, ByVal nomeGrupo As String) As Shape
    Dim s As Shape

    On Error GoTo Falha

    For Each s In shapeRaiz.Shapes.All
        If s.Type = cdrGroupShape Then
            If NomesIguais(s.Name, nomeGrupo) Then
                Set ObterGrupoPorNome = s
                Exit Function
            End If
        End If
    Next s

    Exit Function

Falha:
    Set ObterGrupoPorNome = Nothing
End Function

Public Function BuscarShapePorNomeRecursivo(ByVal raiz As Shape, ByVal nomeBuscado As String) As Shape
    Dim s As Shape
    Dim encontrado As Shape

    On Error GoTo Falha

    If raiz Is Nothing Then Exit Function

    If NomesIguais(raiz.Name, nomeBuscado) Then
        Set BuscarShapePorNomeRecursivo = raiz
        Exit Function
    End If

    If Not TemFilhos(raiz) Then Exit Function

    For Each s In raiz.Shapes.All
        If NomesIguais(s.Name, nomeBuscado) Then
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

Private Function NomesIguais(ByVal nome1 As String, ByVal nome2 As String) As Boolean
    NomesIguais = (StrComp(Trim$(nome1), Trim$(nome2), vbTextCompare) = 0)
End Function

