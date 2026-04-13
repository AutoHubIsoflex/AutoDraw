Attribute VB_Name = "modDescricao"
' modDescricao
Option Explicit

Public Function MontarTextoCompleto(ByVal ehMagnetico As Boolean, _
                                     ByVal altura As Double, _
                                     ByVal largura As Double, _
                                     ByVal catalogo As Collection, _
                                     ByVal contadores As Object, _
                                     ByVal medidasAcessorios As Object) As String
    Dim texto As String
    texto = MontarTextoPrincipal(ehMagnetico, altura, largura)

    Dim secaoAcessorios As String
    secaoAcessorios = MontarTextoAcessorios(catalogo, contadores, medidasAcessorios)

    MontarTextoCompleto = AnexarSecaoAcessorios(texto, secaoAcessorios)
End Function

Private Function MontarTextoPrincipal(ByVal ehMagnetico As Boolean, _
                                       ByVal altura As Double, _
                                       ByVal largura As Double) As String
    If ehMagnetico Then
        MontarTextoPrincipal = "QUADRO BRANCO MAGNÉTICO" & vbCrLf & _
                               "PARA ESCRITA COM IMPRESSÃO " & vbCrLf & _
                               "DIGITAL UV. E LAMINAÇÃO PYT" & vbCrLf & _
                               "MED " & altura & "x" & largura & "MM - QPMM"
    Else
        MontarTextoPrincipal = "QUADRO BRANCO PARA ESCRITA" & vbCrLf & _
                               "COM IMPRESSÃO DIGITAL UV. E" & vbCrLf & _
                               "LAMINAÇÃO PYT MED " & altura & "x" & largura & "MM" & vbCrLf & _
                               "- QPMS"
    End If
End Function

Private Function MontarTextoAcessorios(ByVal catalogo As Collection, _
                                        ByVal contadores As Object, _
                                        ByVal medidasAcessorios As Object) As String
    Dim item As Variant
    Dim nomeShape As String
    Dim quantidade As Long
    Dim texto As String
    Dim outputCode As String

    texto = ""

    For Each item In catalogo
        nomeShape = CStr(item("ShapeName"))
        quantidade = CLng(contadores(nomeShape))
        If quantidade > 0 Then
            outputCode = ResolverOutputCode(nomeShape, CStr(item("OutputCode")), medidasAcessorios)
            If EhCavaleteMetalon3(nomeShape) Then
                texto = texto & "- " & outputCode & vbCrLf
            Else
                texto = texto & "- " & quantidade & " " & outputCode & vbCrLf
            End If
        End If
    Next item

    MontarTextoAcessorios = texto
End Function

Private Function AnexarSecaoAcessorios(ByVal textoPrincipal As String, _
                                        ByVal textoAcessorios As String) As String
    If textoAcessorios = "" Then
        AnexarSecaoAcessorios = textoPrincipal
    Else
        AnexarSecaoAcessorios = textoPrincipal & vbCrLf & vbCrLf & _
                                "ACESSÓRIOS:" & vbCrLf & vbCrLf & _
                                textoAcessorios
    End If
End Function

Private Function EhCavaleteMetalon3(ByVal nomeShape As String) As Boolean
    EhCavaleteMetalon3 = (Left$(UCase$(nomeShape), 18) = "CAVALETE-METALON3-")
End Function

Private Function ResolverOutputCode(ByVal nomeShape As String, _
                                    ByVal outputCode As String, _
                                    ByVal medidasAcessorios As Object) As String
    If UCase$(nomeShape) = "TESTEIRA-MACRO" Then
        If Not medidasAcessorios Is Nothing Then
            If medidasAcessorios.Exists(nomeShape) Then
                ResolverOutputCode = Replace(outputCode, "ALTXLARGURA", CStr(medidasAcessorios(nomeShape)))
                Exit Function
            End If
        End If
    End If

    ResolverOutputCode = outputCode
End Function






