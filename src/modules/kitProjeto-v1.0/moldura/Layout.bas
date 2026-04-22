Attribute VB_Name = "Layout"
Option Explicit

' ==============================================================================
' LAYOUT / GEOMETRIA
' ==============================================================================

Public Sub PosicionarTubosBasicos( _
    ByVal retanguloBase As Shape, _
    ByVal cantSupDir As Shape, ByVal cantSupEsq As Shape, ByVal cantInfDir As Shape, _
    ByVal tuboDir As Shape, ByVal tuboSup As Shape, ByVal tuboEsq As Shape, ByVal tuboInf As Shape)

    tuboDir.LeftX = cantSupDir.RightX - tuboDir.SizeWidth
    tuboDir.CenterY = retanguloBase.CenterY

    tuboSup.CenterX = retanguloBase.CenterX
    tuboSup.TopY = cantSupDir.TopY

    tuboEsq.LeftX = cantSupEsq.LeftX
    tuboEsq.CenterY = retanguloBase.CenterY

    tuboInf.CenterX = retanguloBase.CenterX
    tuboInf.BottomY = cantInfDir.BottomY

End Sub

Public Sub PosicionarTubosEconomy( _
    ByVal retanguloBase As Shape, _
    ByVal cantSupDir As Shape, ByVal cantSupEsq As Shape, ByVal cantInfDir As Shape, _
    ByVal refTuboDir As Shape, ByVal refTuboSup As Shape, _
    ByVal refTuboEsq As Shape, ByVal refTuboInf As Shape, _
    ByVal tuboDir As Shape, ByVal tuboSup As Shape, ByVal tuboEsq As Shape, ByVal tuboInf As Shape)

    If Not refTuboDir Is Nothing Then
        tuboDir.CenterX = refTuboDir.CenterX
        tuboDir.CenterY = retanguloBase.CenterY
    Else
        tuboDir.LeftX = cantSupDir.RightX - tuboDir.SizeWidth
        tuboDir.CenterY = retanguloBase.CenterY
    End If

    tuboSup.CenterX = retanguloBase.CenterX
    If Not refTuboSup Is Nothing Then
        tuboSup.CenterY = refTuboSup.CenterY
    Else
        tuboSup.TopY = cantSupDir.TopY
    End If

    If Not refTuboEsq Is Nothing Then
        tuboEsq.CenterX = refTuboEsq.CenterX
        tuboEsq.CenterY = retanguloBase.CenterY
    Else
        tuboEsq.LeftX = cantSupEsq.LeftX
        tuboEsq.CenterY = retanguloBase.CenterY
    End If

    tuboInf.CenterX = retanguloBase.CenterX
    If Not refTuboInf Is Nothing Then
        tuboInf.CenterY = refTuboInf.CenterY
    Else
        tuboInf.BottomY = cantInfDir.BottomY
    End If

End Sub

Public Sub PosicionarAlhetas( _
    ByVal tuboSup As Shape, ByVal tuboInf As Shape, ByVal deslocamentoHorizontal As Double, _
    ByVal alhetaInfDir As Shape, ByVal alhetaInfEsq As Shape, _
    ByVal alhetaSupDir As Shape, ByVal alhetaSupEsq As Shape)

    alhetaInfEsq.LeftX = tuboInf.LeftX + deslocamentoHorizontal
    alhetaInfEsq.TopY = tuboInf.BottomY

    alhetaInfDir.LeftX = tuboInf.RightX - alhetaInfDir.SizeWidth - deslocamentoHorizontal
    alhetaInfDir.TopY = tuboInf.BottomY

    alhetaSupEsq.LeftX = tuboSup.LeftX + deslocamentoHorizontal
    alhetaSupEsq.BottomY = tuboSup.TopY

    alhetaSupDir.LeftX = tuboSup.RightX - alhetaSupDir.SizeWidth - deslocamentoHorizontal
    alhetaSupDir.BottomY = tuboSup.TopY

End Sub

Public Sub PosicionarCantoneiras(ByVal base As Shape, ByVal offset As Double, _
    ByVal cantSupDir As Shape, ByVal cantSupEsq As Shape, _
    ByVal cantInfEsq As Shape, ByVal cantInfDir As Shape)

    cantSupDir.LeftX = base.RightX - cantSupDir.SizeWidth + offset
    cantSupDir.TopY = base.TopY + offset

    cantSupEsq.LeftX = base.LeftX - offset
    cantSupEsq.TopY = base.TopY + offset

    cantInfEsq.LeftX = base.LeftX - offset
    cantInfEsq.TopY = base.BottomY + cantInfEsq.SizeHeight - offset

    cantInfDir.LeftX = base.RightX - cantInfDir.SizeWidth + offset
    cantInfDir.TopY = base.BottomY + cantInfDir.SizeHeight - offset

End Sub

Public Sub PosicionarCantoneirasEconomy(ByVal base As Shape, ByVal desl As Double, _
    ByVal cantSupDir As Shape, ByVal cantSupEsq As Shape, _
    ByVal cantInfEsq As Shape, ByVal cantInfDir As Shape)

    cantSupDir.LeftX = base.RightX - cantSupDir.SizeWidth + desl
    cantSupDir.TopY = base.TopY + desl

    cantSupEsq.LeftX = base.LeftX - desl
    cantSupEsq.TopY = base.TopY + desl

    cantInfEsq.LeftX = base.LeftX - desl
    cantInfEsq.TopY = base.BottomY + cantInfEsq.SizeHeight - desl

    cantInfDir.LeftX = base.RightX - cantInfDir.SizeWidth + desl
    cantInfDir.TopY = base.BottomY + cantInfDir.SizeHeight - desl

End Sub

Public Sub AjustarTubosEntreCentros( _
    ByVal cantSupDir As Shape, ByVal cantSupEsq As Shape, _
    ByVal cantInfEsq As Shape, ByVal cantInfDir As Shape, _
    ByVal tuboDir As Shape, ByVal tuboSup As Shape, _
    ByVal tuboEsq As Shape, ByVal tuboInf As Shape)

    Dim largura As Double
    Dim altura As Double

    largura = cantInfDir.CenterX - cantInfEsq.CenterX
    tuboInf.SetSize largura, tuboInf.SizeHeight
    tuboInf.CenterX = (cantInfEsq.CenterX + cantInfDir.CenterX) / 2

    altura = cantSupEsq.CenterY - cantInfEsq.CenterY
    tuboEsq.SetSize tuboEsq.SizeWidth, altura
    tuboEsq.CenterY = (cantSupEsq.CenterY + cantInfEsq.CenterY) / 2

    altura = cantSupDir.CenterY - cantInfDir.CenterY
    tuboDir.SetSize tuboDir.SizeWidth, altura
    tuboDir.CenterY = (cantSupDir.CenterY + cantInfDir.CenterY) / 2

    largura = cantSupDir.CenterX - cantSupEsq.CenterX
    tuboSup.SetSize largura, tuboSup.SizeHeight
    tuboSup.CenterX = (cantSupEsq.CenterX + cantSupDir.CenterX) / 2

End Sub

