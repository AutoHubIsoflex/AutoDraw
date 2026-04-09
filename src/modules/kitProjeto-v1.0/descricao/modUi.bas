Attribute VB_Name = "modUi"
' modUI
Option Explicit

Public Function SolicitarTipoQuadro() As tipoQuadro
    On Error GoTo Falha

    frmTipoQuadro.TipoSelecionado = ""
    frmTipoQuadro.Show vbModal

    Dim selecionado As String
    selecionado = frmTipoQuadro.TipoSelecionado
    Unload frmTipoQuadro

    Select Case selecionado
        Case "QPMS-P": SolicitarTipoQuadro = tqQPMS_P
        Case "QPMM-P": SolicitarTipoQuadro = tqQPMM_P
        Case Else:     SolicitarTipoQuadro = -1
    End Select

    Exit Function
Falha:
    MsgBox "UserForm 'frmTipoQuadro' não encontrado ou inválido.", vbCritical
    SolicitarTipoQuadro = -1
End Function

Public Function ConfirmarCompatibilidade(ByVal tipo As tipoQuadro, _
                                          ByVal ehMG As Boolean, _
                                          ByVal ehAD As Boolean) As Boolean
    Dim resposta As VbMsgBoxResult

    If tipo = tqQPMS_P And ehMG Then
        resposta = MsgBox("Acessórios magnéticos encontrados, deseja continuar?", _
                          vbQuestion + vbYesNo, "Validação de acessórios")
        If resposta = vbNo Then
            ConfirmarCompatibilidade = False
            Exit Function
        End If
    End If

    If tipo = tqQPMM_P And ehAD Then
        resposta = MsgBox("Acessórios adesivos encontrados, deseja continuar?", _
                          vbQuestion + vbYesNo, "Validação de acessórios")
        If resposta = vbNo Then
            ConfirmarCompatibilidade = False
            Exit Function
        End If
    End If

    ConfirmarCompatibilidade = True
End Function


