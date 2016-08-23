Attribute VB_Name = "modUserForm"
Option Explicit



Sub MakeUserForm()
     
    Dim MyUserForm As VBComponent
    Dim NewOptionButton As MSForms.OptionButton
    Dim NewCommandButton1 As MSForms.CommandButton
    Dim NewCommandButton2 As MSForms.CommandButton
    Dim MyComboBox As MSForms.ComboBox
    Dim N, X As Integer, MaxWidth As Long
     
     '//First, check the form doesn't already exist
    For N = 1 To ActiveWorkbook.VBProject.VBComponents.Count
        If ActiveWorkbook.VBProject.VBComponents(N).name = "NewForm" Then
            ShowForm
            Exit Sub
        Else
        End If
    Next N
     
     '//Make a userform
    Set MyUserForm = ActiveWorkbook.VBProject _
    .VBComponents.Add(vbext_ct_MSForm)
    With MyUserForm
        .Properties("Height") = 100
        .Properties("Width") = 200
        On Error Resume Next
        .name = "NewForm"
        .Properties("Caption") = "Here is your user form"
    End With
     
     '//Add a Cancel button to the form
    Set NewCommandButton1 = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With NewCommandButton1
        .Caption = "Cancel"
        .Height = 18
        .Width = 44
        .Left = MaxWidth + 147
        .Top = 6
    End With
     
     '//Add an OK button to the form
    Set NewCommandButton2 = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With NewCommandButton2
        .Caption = "OK"
        .Height = 18
        .Width = 44
        .Left = MaxWidth + 147
        .Top = 28
    End With
     
     '//Add code on the form for the CommandButtons
    With MyUserForm.CodeModule
        X = .countOfLines
        Dim i As Long
        For i = 1 To UBound(arrayModuloForm)
            .InsertLines X + i, arrayModuloForm(i)
        Next i
    End With
     
     '//Add a combo box on the form
    Set MyComboBox = MyUserForm.Designer.Controls.Add("Forms.ComboBox.1")
    With MyComboBox
        .name = "Combo1"
        .Left = 10
        .Top = 10
        .Height = 16
        .Width = 100
    End With
     
    ShowForm
End Sub
 
Sub ShowForm()
    NewForm.Show
End Sub

