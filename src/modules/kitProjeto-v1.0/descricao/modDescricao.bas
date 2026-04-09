Attribute VB_Name = "modDescricao"
' modDescricao
Option Explicit

Public Function MontarTextoCompleto(ByVal ehMagnetico As Boolean, _
                                     ByVal altura As Double, _
                                     ByVal largura As Double, _
                                     ByVal catalogo As Collection, _
                                     ByVal contadores As Object) As String
    Dim texto As String
    texto = MontarTextoPrincipal(ehMagnetico, altura, largura)

    Dim secaoAcessorios As String
    secaoAcessorios = MontarTextoAcessorios(catalogo, contadores)

    MontarTextoCompleto = AnexarSecaoAcessorios(texto, secaoAcessorios)
End Function

Private Function MontarTextoPrincipal(ByVal ehMagnetico As Boolean, _
                                       ByVal altura As Double, _
                                       ByVal largura As Double) As String
    If ehMagnetico Then
        MontarTextoPrincipal = "QUADRO BRANCO MAGNÉTICO" & vbCrLf & _
                               "PARA ESCRITA COM IMPRESSÃO " & vbCrLf & _
                               "DIGITAL UV. E LAMINAÇÃO PYT" & vbCrLf & _
                               "MED " & altura & "x" & largura & " - QPMM"
    Else
        MontarTextoPrincipal = "QUADRO BRANCO PARA ESCRITA" & vbCrLf & _
                               "COM IMPRESSÃO DIGITAL UV. E" & vbCrLf & _
                               "LAMINAÇÃO PYT MED " & altura & "x" & largura & vbCrLf & _
                               "- QPMS"
    End If
End Function

Private Function MontarTextoAcessorios(ByVal catalogo As Collection, _
                                        ByVal contadores As Object) As String
    Dim item As Variant
    Dim nomeShape As String
    Dim quantidade As Long
    Dim texto As String

    texto = ""

    For Each item In catalogo
        nomeShape = CStr(item("ShapeName"))
        quantidade = CLng(contadores(nomeShape))
        If quantidade > 0 Then
            texto = texto & "- " & quantidade & " " & CStr(item("OutputCode")) & vbCrLf
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


