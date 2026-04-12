Attribute VB_Name = "modCatalogo"
' modCatalogo
Option Explicit

Public Enum tipoQuadro
    tqQPMS_P = 0
    tqQPMM_P = 1
End Enum

Public Const COMPAT_AD     As String = "AD"
Public Const COMPAT_MG     As String = "MG"
Public Const COMPAT_NEUTRO As String = "NEUTRO"

Private Function ObterDefinicoesBrutas() As Variant
    ObterDefinicoesBrutas = Array( _
        "KSIP-A3-AD-MACRO|KSIP-A3-AD|AD", _
        "KSIP-A3-MG-MACRO|KSIP-A3-MG|MG", _
        "KSIR-A3-AD-MACRO|KSIR-A3-AD|AD", _
        "KSIR-A3-MG-MACRO|KSIR-A3-MG|MG", _
        "KSIR-A4-AD-MACRO|KSIR-A4-AD|AD", _
        "KSIP-A4-AD-MACRO|KSIP-A4-AD|AD", _
        "KSIR-A4-MG-MACRO|KSIR-A4-MG|MG", _
        "KSIP-A4-MG-MACRO|KSIP-A4-MG|MG", _
        "ESC-A4-CZ-MACRO|ESC-A4-CZ|NEUTRO", _
        "ESC-A4-AM-MACRO|ESC-A4-AM|NEUTRO", _
        "ESC-A4-AZ-MACRO|ESC-A4-AZ|NEUTRO", _
        "ESC-A4-VD-MACRO|ESC-A4-VD|NEUTRO", _
        "ESC-A4-VM-MACRO|ESC-A4-VM|NEUTRO", _
        "ESC-A4-PT-MACRO|ESC-A4-PT|NEUTRO", _
        "BASE-ESC-A4-MACRO|BASE-ESC-A4|NEUTRO", _
        "CAVALETE-METALON3-BR|CAVALETE METALON 3 BR|NEUTRO", _
        "CAVALETE-METALON3-CZ|CAVALETE METALON 3 CZ|NEUTRO", _
        "CAVALETE-METALON3-PT|CAVALETE METALON 3 PT|NEUTRO" _
    )
End Function

Public Function CriarCatalogoAcessorios() As Collection
    Dim catalogo As New Collection
    Dim def As Variant
    Dim partes() As String

    For Each def In ObterDefinicoesBrutas()
        partes = Split(CStr(def), "|")
        catalogo.Add CriarDefinicaoAcessorio(partes(0), partes(1), partes(2))
    Next def

    Set CriarCatalogoAcessorios = catalogo
End Function

Public Function CriarIndiceAcessorios(ByVal catalogo As Collection) As Object
    Dim indice As Object
    Dim item As Variant

    Set indice = CreateObject("Scripting.Dictionary")
    For Each item In catalogo
        indice.Add CStr(item("ShapeName")), item
    Next item

    Set CriarIndiceAcessorios = indice
End Function

Private Function CriarDefinicaoAcessorio(ByVal shapeName As String, _
                                          ByVal outputCode As String, _
                                          ByVal compat As String) As Object
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item.Add "ShapeName", UCase$(shapeName)
    item.Add "OutputCode", outputCode
    item.Add "Compat", compat

    Set CriarDefinicaoAcessorio = item
End Function







