Attribute VB_Name = "modAuxiliar"
Public Function SelecionarRange() As Range
On Error GoTo HandleError
    Dim ThisRng As Range
    Set ThisRng = Application.InputBox("Select a range", "Get Range", Type:=8)
    Set SelecionarRange = ThisRng
    Exit Function
HandleError:
    Debug.Print "Range selecionado � inv�lido"
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
 
    'Acentos e caracteres especiais que ser�o buscados na string
    'Voc� pode definir outros caracteres nessa vari�vel, mas
    ' precisar� tamb�m colocar a letra correspondente em codiB
    codiA = "�����������������������������������������������/ -$()%"
     
    'Letras correspondentes para substitui��o
    codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN___S__P"
     
    'Armazena em temp a string recebida
    temp = caract
     
    'Loop que ir� de andar� a string letra a letra
    For i = 1 To Len(temp)
     
        'InStr buscar� se a letra indice i de temp pertence a
        ' codiA e se existir retornar� a posi��o dela
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

Public Function Array2DTranspose(avValues As Variant) As Variant
    Dim lThisCol As Long, lThisRow As Long
    Dim lUb2 As Long, lLb2 As Long
    Dim lUb1 As Long, lLb1 As Long
    Dim avTransposed As Variant

    If IsArray(avValues) Then
        On Error GoTo ErrFailed
        lUb2 = UBound(avValues, 2)
        lLb2 = LBound(avValues, 2)
        lUb1 = UBound(avValues, 1)
        lLb1 = LBound(avValues, 1)

        ReDim avTransposed(lLb2 To lUb2, lLb1 To lUb1)
        For lThisCol = lLb1 To lUb1
            For lThisRow = lLb2 To lUb2
                avTransposed(lThisRow, lThisCol) = avValues(lThisCol, lThisRow)
            Next
        Next
    End If

    Array2DTranspose = avTransposed
    Exit Function

ErrFailed:
    Debug.Print Err.Description
    Debug.Assert False
    Array2DTranspose = Empty
    Exit Function
    Resume
End Function
