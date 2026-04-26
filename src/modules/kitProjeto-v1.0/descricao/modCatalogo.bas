Attribute VB_Name = "modCatalogo"
' modCatalogo
Option Explicit

Public Enum tipoQuadro
    tqQPMS_P = 0
    tqQPMM_P = 1
    tqQBTA = 2
End Enum

Public Const COMPAT_AD     As String = "AD"
Public Const COMPAT_MG     As String = "MG"
Public Const COMPAT_NEUTRO As String = "NEUTRO"
Public Const SHAPE_REFORCO_ALUMINIO_AUTO As String = "REFORCO-ALUMINIO-AUTO"

Private Function ObterDefinicoesBrutas() As Variant
    Dim defs As Collection
    Set defs = New Collection

    defs.Add "KSIP-A3-AD-MACRO|KSIP A3 AD|AD"
    defs.Add "KSIP-A3-MG-MACRO|KSIP A3 MG|MG"
    defs.Add "KSIR-A3-AD-MACRO|KSIR A3 AD|AD"
    defs.Add "KSIR-A3-MG-MACRO|KSIR A3 MG|MG"
    defs.Add "KSIR-A4-AD-MACRO|KSIR A4 AD|AD"
    defs.Add "KSIP-A4-AD-MACRO|KSIP A4 AD|AD"
    defs.Add "KSIR-A4-MG-MACRO|KSIR A4 MG|MG"
    defs.Add "KSIP-A4-MG-MACRO|KSIP A4 MG|MG"
    defs.Add "ESC-A4-CZ-MACRO|ESC A4 CZ|NEUTRO"
    defs.Add "ESC-A4-AM-MACRO|ESC A4 AM|NEUTRO"
    defs.Add "ESC-A4-AZ-MACRO|ESC A4 AZ|NEUTRO"
    defs.Add "ESC-A4-VD-MACRO|ESC A4 VD|NEUTRO"
    defs.Add "ESC-A4-VM-MACRO|ESC A4 VM|NEUTRO"
    defs.Add "ESC-A4-PT-MACRO|ESC A4 PT|NEUTRO"
    defs.Add "PTI-MACRO|PTI|NEUTRO"
    defs.Add "PTC-MACRO|PTC|NEUTRO"
    defs.Add "PTL-MACRO|PTL|NEUTRO"
    defs.Add "DAVN-MACRO|DAVN-TIPO-MED ALTXLARGURAMM|NEUTRO"
    defs.Add "TESTEIRA-MACRO|TEST PS MED ALTXLARGURAMM|NEUTRO"
    defs.Add "BASE-ESC-A4-MACRO|BASE ESC A4|NEUTRO"
    defs.Add "BASE-BIG-ISOLEAN-MACRO|BASE BIG ISOLEAN A4|NEUTRO"
    defs.Add "BIG-ISOLEAN-AM-MACRO|BIG ISOLEAN AM A4|NEUTRO"
    defs.Add "BIG-ISOLEAN-AZ-MACRO|BIG ISOLEAN AZ A4|NEUTRO"
    defs.Add "BIG-ISOLEAN-VM-MACRO|BIG ISOLEAN VM A4|NEUTRO"
    defs.Add "BIG-ISOLEAN-VD-MACRO|BIG ISOLEAN VD A4|NEUTRO"
    defs.Add "BIG-ISOLEAN-CZ-MACRO|BIG ISOLEAN CZ A4|NEUTRO"
    defs.Add "BIG-ISOLEAN-CR-MACRO|BIG ISOLEAN CR A4|NEUTRO"
    defs.Add "CAVALETE-METALON3-BR|CAVALETE METALON 3 BR|NEUTRO"
    defs.Add "CAVALETE-METALON3-CZ|CAVALETE METALON 3 CZ|NEUTRO"
    defs.Add "CAVALETE-METALON3-PT|CAVALETE METALON 3 PT|NEUTRO"
    defs.Add "KSIR-A5-AD-MACRO|KSIR A5 AD|AD"
    defs.Add "KSIR-A5-MG-MACRO|KSIR A5 MG|MG"
    defs.Add "KSIP-A5-AD-MACRO|KSIP A5 AD|AD"
    defs.Add "KSIP-A5-MG-MACRO|KSIP A5 MG|MG"
    defs.Add "KSVR-A4-AD-MACRO|KSVR A4 AD|AD"
    defs.Add "KSVR-A4-MG-MACRO|KSVR A4 MG|MG"
    defs.Add "KSVP-A4-AD-MACRO|KSVP A4 AD|AD"
    defs.Add "KSVP-A4-MG-MACRO|KSVP A4 MG|MG"
    defs.Add "KSIR-A2-AD-MACRO|KSIR A2 AD|AD"
    defs.Add "KSIR-A2-MG-MACRO|KSIR A2 MG|MG"
    defs.Add "KSIP-A2-AD-MACRO|KSIP A2 AD|AD"
    defs.Add "KSIP-A2-MG-MACRO|KSIP A2 MG|MG"
    defs.Add "KSIR-A6-AD-MACRO|KSIR A6 AD|AD"
    defs.Add "KSIR-A6-MG-MACRO|KSIR A6 MG|MG"
    defs.Add "KSIP-A6-AD-MACRO|KSIP A6 AD|AD"
    defs.Add "KSIP-A6-MG-MACRO|KSIP A6 MG|MG"
    defs.Add "REGUA-BENSON-FIX-MACRO|RÉGUA BENSON COM 4 GARRAS" & vbCrLf & "FIX/ PARAFUSADA|NEUTRO"
    defs.Add "REGUA-BENSON-MG-MACRO|RÉGUA BENSON COM 4 GARRAS" & vbCrLf & "FIX/ MAGNÉTICA|MG"
    defs.Add SHAPE_REFORCO_ALUMINIO_AUTO & "|PAR DE REFORÇO EM ALUMÍNIO" & vbCrLf & "PARA QUADROS|NEUTRO"
    
    Set ObterDefinicoesBrutas = defs
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




