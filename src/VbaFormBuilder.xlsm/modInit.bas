Attribute VB_Name = "modInit"
Public Sub ImportarCodigo()
    On Error GoTo trataErro
    Call testImport
    MsgBox "Importação realizada com sucesso"
    Exit Sub
trataErro:
    MsgBox "Ocorreu um erro na importação do código. Feche o arquivo sem salvar e tente novamente. Código do erro: " & Err.Number & " - " & Err.Description
End Sub
