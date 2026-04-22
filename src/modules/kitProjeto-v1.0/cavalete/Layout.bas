Attribute VB_Name = "Layout"
Option Explicit

' =========================================================
' LAYOUT / POSICIONAMENTO
' =========================================================

Public Sub PosicionarCavaleteInicial(ByVal cavalete As Shape, ByVal quadro As Shape)
    cavalete.TopY = quadro.TopY
    cavalete.LeftX = quadro.LeftX

    cavalete.LeftX = cavalete.LeftX - MmParaDocumento(DESLOCAMENTO_X_CAVALETE_MM)
    cavalete.TopY = cavalete.TopY + MmParaDocumento(DESLOCAMENTO_Y_CAVALETE_MM)
End Sub

Public Sub PosicionarMaoFrancesa(ByVal maoFrancesa As Shape, ByVal quadro As Shape)
    maoFrancesa.BottomY = quadro.BottomY
    maoFrancesa.BottomY = maoFrancesa.BottomY - MmParaDocumento(DESLOCAMENTO_Y_MAO_FRANCESA_MM)
End Sub

Public Sub EspelharEPosicionarGrupo(ByVal grupo As Shape, ByVal quadro As Shape)
    grupo.Flip cdrFlipHorizontal
    grupo.RightX = quadro.RightX
    grupo.RightX = grupo.RightX + MmParaDocumento(DESLOCAMENTO_X_GRUPO_ESPELHADO_MM)
End Sub

Public Function MmParaDocumento(ByVal valorMm As Double) As Double
    MmParaDocumento = ActiveDocument.ToUnits(valorMm, cdrMillimeter)
End Function

