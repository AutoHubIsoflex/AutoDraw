Attribute VB_Name = "Map"
Option Explicit

' ==============================================================================
' MAPEAMENTO DE PEÇAS
' ==============================================================================

Public Sub MapearPecasBasicas(ByVal grupo As Shape, _
    ByRef cantSupDir As Shape, ByRef cantSupEsq As Shape, _
    ByRef cantInfEsq As Shape, ByRef cantInfDir As Shape, _
    ByRef tuboDir As Shape, ByRef tuboSup As Shape, _
    ByRef tuboEsq As Shape, ByRef tuboInf As Shape)

    Dim s As Shape

    For Each s In grupo.Shapes
        AtribuirPecaComumPorNome s, cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
                                  tuboDir, tuboSup, tuboEsq, tuboInf
    Next s

End Sub

Public Sub MapearPecasEconomy(ByVal grupo As Shape, _
    ByRef cantSupDir As Shape, ByRef cantSupEsq As Shape, _
    ByRef cantInfEsq As Shape, ByRef cantInfDir As Shape, _
    ByRef tuboDir As Shape, ByRef tuboSup As Shape, _
    ByRef tuboEsq As Shape, ByRef tuboInf As Shape, _
    ByRef alhetaInfDir As Shape, ByRef alhetaInfEsq As Shape, _
    ByRef alhetaSupDir As Shape, ByRef alhetaSupEsq As Shape)

    Dim s As Shape

    For Each s In grupo.Shapes
        AtribuirPecaComumPorNome s, cantSupDir, cantSupEsq, cantInfEsq, cantInfDir, _
                                  tuboDir, tuboSup, tuboEsq, tuboInf
        AtribuirAlhetaPorNome s, alhetaInfDir, alhetaInfEsq, alhetaSupDir, alhetaSupEsq
    Next s

End Sub

Public Sub AtribuirPecaComumPorNome( _
    ByVal peca As Shape, _
    ByRef cantSupDir As Shape, ByRef cantSupEsq As Shape, _
    ByRef cantInfEsq As Shape, ByRef cantInfDir As Shape, _
    ByRef tuboDir As Shape, ByRef tuboSup As Shape, _
    ByRef tuboEsq As Shape, ByRef tuboInf As Shape)

    Select Case peca.Name
        Case NOME_CANT_SUP_DIR: Set cantSupDir = peca
        Case NOME_CANT_SUP_ESQ: Set cantSupEsq = peca
        Case NOME_CANT_INF_ESQ: Set cantInfEsq = peca
        Case NOME_CANT_INF_DIR: Set cantInfDir = peca
        Case NOME_TUBO_DIR: Set tuboDir = peca
        Case NOME_TUBO_SUP: Set tuboSup = peca
        Case NOME_TUBO_ESQ: Set tuboEsq = peca
        Case NOME_TUBO_INF: Set tuboInf = peca
    End Select

End Sub

Public Sub AtribuirAlhetaPorNome( _
    ByVal peca As Shape, _
    ByRef alhetaInfDir As Shape, ByRef alhetaInfEsq As Shape, _
    ByRef alhetaSupDir As Shape, ByRef alhetaSupEsq As Shape)

    Select Case peca.Name
        Case NOME_ALHETA_INF_DIR: Set alhetaInfDir = peca
        Case NOME_ALHETA_INF_ESQ: Set alhetaInfEsq = peca
        Case NOME_ALHETA_SUP_DIR: Set alhetaSupDir = peca
        Case NOME_ALHETA_SUP_ESQ: Set alhetaSupEsq = peca
    End Select

End Sub

Public Function BuscarShapePorNome(ByVal grp As Shape, ByVal nome As String) As Shape

    Dim s As Shape
    Dim resultado As Shape

    If grp.Type <> cdrGroupShape Then Exit Function

    For Each s In grp.Shapes

        If s.Name = nome Then
            Set BuscarShapePorNome = s
            Exit Function
        End If

        If s.Type = cdrGroupShape Then
            Set resultado = BuscarShapePorNome(s, nome)
            If Not resultado Is Nothing Then
                Set BuscarShapePorNome = resultado
                Exit Function
            End If
        End If

    Next s

End Function


