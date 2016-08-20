Attribute VB_Name = "modAuxiliar"
Public Function SelecionarRange() As Range
On Error GoTo HandleError
    Dim ThisRng As Range
    Set ThisRng = Application.InputBox("Select a range", "Get Range", Type:=8)
    Set SelecionarRange = ThisRng
    Exit Function
HandleError:
    Debug.Print "Range selecionado é inválido"
    Set SelecionarRange = Nothing
End Function

Public Function IsVarArrayEmpty(anArray As Variant)

Dim i As Integer

On Error Resume Next
    i = UBound(anArray, 1)
If Err.Number = 0 Then
    IsVarArrayEmpty = False
Else
    IsVarArrayEmpty = True
End If

End Function

