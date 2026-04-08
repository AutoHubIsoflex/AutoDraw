VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTipoQuadro 
   Caption         =   "Isoflex"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmTipoQuadro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTipoQuadro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TipoSelecionado As String

Private Sub lblTipo_Click()

End Sub

Private Sub UserForm_Initialize()
    With Me.cboTipoQuadro
        .Clear
        .AddItem "QPMS-P"
        .AddItem "QPMM-P"
        .ListIndex = 0
    End With

    TipoSelecionado = ""
End Sub

Private Sub btnConfirmar_Click()
    If Me.cboTipoQuadro.ListIndex < 0 Then
        MsgBox "Selecione um tipo de quadro.", vbExclamation
        Exit Sub
    End If

    TipoSelecionado = Me.cboTipoQuadro.Value
    Me.Hide
End Sub

Private Sub btnCancelar_Click()
    TipoSelecionado = ""
    Me.Hide
End Sub

Private Sub cboTipoQuadro_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnConfirmar_Click
End Sub
