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

Public Function RemoveAcentos(ByVal caract As String) As String
 
    'Acentos e caracteres especiais que serão buscados na string
    'Você pode definir outros caracteres nessa variável, mas
    ' precisará também colocar a letra correspondente em codiB
    codiA = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ/ -$()"
     
    'Letras correspondentes para substituição
    codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN___S__"
     
    'Armazena em temp a string recebida
    temp = caract
     
    'Loop que irá de andará a string letra a letra
    For i = 1 To Len(temp)
     
        'InStr buscará se a letra indice i de temp pertence a
        ' codiA e se existir retornará a posição dela
        p = InStr(codiA, Mid(temp, i, 1))
         
        'Substitui a letra de indice i em codiA pela sua
        ' correspondente em codiB
        If p > 0 Then Mid(temp, i, 1) = Mid(codiB, p, 1)
    Next
     
    'Retorna a nova string
    RemoveAcentos = temp
     
End Function


Public Function CleanString(text As String) As String
    Dim output As String
    Dim c 'since char type does not exist in vba, we have to use variant type.
    For i = 1 To Len(text)
        c = Mid(text, i, 1) 'Select the character at the i position
        If (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Or (c >= "A" And c <= "Z") Then
            output = output & c 'add the character to your output.
        Else
            output = output & "_" 'add the replacement character (space) to your output
        End If
    Next
    CleanString = output
End Function
