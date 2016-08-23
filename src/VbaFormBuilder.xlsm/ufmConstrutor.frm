VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmConstrutor 
   Caption         =   "Construtor de Formulários"
   ClientHeight    =   4212
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7416
   OleObjectBlob   =   "ufmConstrutor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufmConstrutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private controles()
Private arrayModuloForm(1 To 266)
Private arrayModuloFuncaoLimpaControles(1 To 3)
Private arrayModuloFuncaoControlDataType(1 To 6)
Private arrayModuloFuncaoSetValues(1 To 7)
Private arrayModuloFuncaoGetValues(1 To 6)
Private nomeCampoChavePrimaria As String
Private countOfLines As Long
Private Const colunaCampo As Integer = 1
Private Const colunaControle As Integer = 2
Private Const colunaRequerido As Integer = 3
Private Const colunaEchave As Integer = 4

Public Sub DefineControles(ByRef pControles())
     controles = pControles
End Sub

Private Sub Init()
    arrayModuloForm(1) = "Public IsCancelled As Boolean"
    arrayModuloForm(2) = "Private cls[NOME_FORM] As [NOME_FORM]"
    arrayModuloForm(3) = "Private modoEdicao As Boolean"
    arrayModuloForm(4) = ""
    arrayModuloForm(5) = "Private Sub AlteraModo(ByVal Edicao As Boolean)"
    arrayModuloForm(6) = "    Dim ctl As MSForms.Control"
    arrayModuloForm(7) = "    'controles de input"
    arrayModuloForm(8) = "    For Each ctl In Me.Controls"
    arrayModuloForm(9) = "        If IsInputControl(ctl) Then"
    arrayModuloForm(10) = "           ctl.Enabled = Edicao"
    arrayModuloForm(11) = "        End If"
    arrayModuloForm(12) = "    Next"
    arrayModuloForm(13) = "    "
    arrayModuloForm(14) = "    'excessão"
    arrayModuloForm(15) = "    txt[CHAVE_PRIMARIA].Enabled = False"
    arrayModuloForm(16) = "    "
    arrayModuloForm(17) = "    'botoes de navegacao"
    arrayModuloForm(18) = "    btnOk.Enabled = Edicao"
    arrayModuloForm(19) = "    btnCancel.Enabled = Edicao"
    arrayModuloForm(20) = "    btnPrimeiro.Enabled = Not Edicao"
    arrayModuloForm(21) = "    btnAnterior.Enabled = Not Edicao"
    arrayModuloForm(22) = "    btnProximo.Enabled = Not Edicao"
    arrayModuloForm(23) = "    btnUtimo.Enabled = Not Edicao"
    arrayModuloForm(24) = "    'os options buttons de operacao"
    arrayModuloForm(25) = "    optAlterar.Enabled = Not Edicao"
    arrayModuloForm(26) = "    optExcluir.Enabled = Not Edicao"
    arrayModuloForm(27) = "    optNovo.Enabled = Not Edicao"
    arrayModuloForm(28) = "   "
    arrayModuloForm(29) = "    If Not Edicao Then"
    arrayModuloForm(30) = "        optAlterar.Value = False"
    arrayModuloForm(31) = "        optExcluir.Value = False"
    arrayModuloForm(32) = "        optNovo.Value = False"
    arrayModuloForm(33) = "        lblStatus.Caption = """""
    arrayModuloForm(34) = "    End If"
    arrayModuloForm(35) = "    "
    arrayModuloForm(36) = "    modoEdicao = Edicao"
    arrayModuloForm(37) = "End Sub"
    arrayModuloForm(38) = ""
    arrayModuloForm(39) = "Private Sub btnAnterior_Click()"
    arrayModuloForm(40) = "    If cls[NOME_FORM].MovePrevious Then Call SetValues(cls[NOME_FORM])"
    arrayModuloForm(41) = "End Sub"
    arrayModuloForm(42) = ""
    arrayModuloForm(43) = "Private Sub btnPesquisar_Click()"
    arrayModuloForm(44) = "    ufm[NOME_FORM]Pesquisa.Show"
    arrayModuloForm(45) = "End Sub"
    arrayModuloForm(46) = ""
    arrayModuloForm(47) = "Private Sub btnPrimeiro_Click()"
    arrayModuloForm(48) = "    cls[NOME_FORM].MoveFirst"
    arrayModuloForm(49) = "    Call SetValues(cls[NOME_FORM])"
    arrayModuloForm(50) = "End Sub"
    arrayModuloForm(51) = ""
    arrayModuloForm(52) = "Private Sub btnProximo_Click()"
    arrayModuloForm(53) = "    If cls[NOME_FORM].MoveNext Then Call SetValues(cls[NOME_FORM])"
    arrayModuloForm(54) = "End Sub"
    arrayModuloForm(55) = ""
    arrayModuloForm(56) = "Private Sub btnUtimo_Click()"
    arrayModuloForm(57) = "    cls[NOME_FORM].MoveLast"
    arrayModuloForm(58) = "    Call SetValues(cls[NOME_FORM])"
    arrayModuloForm(59) = "End Sub"
    arrayModuloForm(60) = ""
    arrayModuloForm(61) = "Private Sub optAlterar_Click()"
    arrayModuloForm(62) = "    AlteraModo Edicao:=True"
    arrayModuloForm(63) = "End Sub"
    arrayModuloForm(64) = ""
    arrayModuloForm(65) = "Private Sub optExcluir_Click()"
    arrayModuloForm(66) = "    lblStatus.Caption = ""Modo de exclusão"""
    arrayModuloForm(67) = "    AlteraModo Edicao:=True"
    arrayModuloForm(68) = "End Sub"
    arrayModuloForm(69) = ""
    arrayModuloForm(70) = "Private Sub optNovo_Click()"
    arrayModuloForm(71) = "    AlteraModo Edicao:=True"
    arrayModuloForm(72) = "    Call LimpaControles"
    arrayModuloForm(73) = "    cls[NOME_FORM].AddNew"
    arrayModuloForm(74) = "End Sub"
    arrayModuloForm(75) = ""
    arrayModuloForm(76) = "Private Sub UserForm_Initialize()"
    arrayModuloForm(77) = "    IsCancelled = True"
    arrayModuloForm(78) = "    "
    arrayModuloForm(79) = "    AlteraModo Edicao:=False"
    arrayModuloForm(80) = "End Sub"
    arrayModuloForm(81) = ""
    arrayModuloForm(82) = "Private Sub btnCancel_Click()"
    arrayModuloForm(83) = "    AlteraModo Edicao:=False"
    arrayModuloForm(84) = "    cls[NOME_FORM].MovePrevious"
    arrayModuloForm(85) = "    Call SetValues(cls[NOME_FORM])"
    arrayModuloForm(86) = "    'Me.Hide"
    arrayModuloForm(87) = "End Sub"
    arrayModuloForm(88) = ""
    arrayModuloForm(89) = "Private Sub btnOK_Click()"
    arrayModuloForm(90) = "    If optExcluir.Value Then"
    arrayModuloForm(91) = "        If MsgBox(""Deseja realmente excluir este registro?"", vbYesNo, ""Aviso de Exclusão"") = vbYes Then"
    arrayModuloForm(92) = "            cls[NOME_FORM].Delete"
    arrayModuloForm(93) = "            AlteraModo Edicao:=False"
    arrayModuloForm(94) = "            cls[NOME_FORM].MoveFirst"
    arrayModuloForm(95) = "            Call SetValues(cls[NOME_FORM])"
    arrayModuloForm(96) = "        End If"
    arrayModuloForm(97) = "    ElseIf IsInputOk Then"
    arrayModuloForm(98) = "        IsCancelled = False"
    arrayModuloForm(99) = "        Call GetValues(cls[NOME_FORM])"
    arrayModuloForm(100) = "        If cls[NOME_FORM].Update Then"
    arrayModuloForm(101) = "            AlteraModo Edicao:=False"
    arrayModuloForm(102) = "            cls[NOME_FORM].MoveFirst"
    arrayModuloForm(103) = "            Call SetValues(cls[NOME_FORM])"
    arrayModuloForm(104) = "        End If"
    arrayModuloForm(105) = "        'Me.Hide"
    arrayModuloForm(106) = "    End If"
    arrayModuloForm(107) = "End Sub"
    arrayModuloForm(108) = ""
    arrayModuloForm(109) = "Private Function IsInputOk() As Boolean"
    arrayModuloForm(110) = "Dim ctl As MSForms.Control"
    arrayModuloForm(111) = "Dim strMessage As String"
    arrayModuloForm(112) = "    IsInputOk = False"
    arrayModuloForm(113) = "    For Each ctl In Me.Controls"
    arrayModuloForm(114) = "        If IsInputControl(ctl) Then"
    arrayModuloForm(115) = "            If IsRequired(ctl) Then"
    arrayModuloForm(116) = "                If Not HasValue(ctl) Then"
    arrayModuloForm(117) = "                    strMessage = ControlName(ctl) & "" é obrigatório"""
    arrayModuloForm(118) = "                End If"
    arrayModuloForm(119) = "            End If"
    arrayModuloForm(120) = "            If Not IsCorrectType(ctl) Then"
    arrayModuloForm(121) = "                strMessage = ControlName(ctl) & "" é inválido"""
    arrayModuloForm(122) = "            End If"
    arrayModuloForm(123) = "        End If"
    arrayModuloForm(124) = "        If Len(strMessage) > 0 Then"
    arrayModuloForm(125) = "            ctl.SetFocus"
    arrayModuloForm(126) = "            GoTo HandleMessage"
    arrayModuloForm(127) = "        End If"
    arrayModuloForm(128) = "    Next"
    arrayModuloForm(129) = "    IsInputOk = True"
    arrayModuloForm(130) = "HandleExit:"
    arrayModuloForm(131) = "    Exit Function"
    arrayModuloForm(132) = "HandleMessage:"
    arrayModuloForm(133) = "    MsgBox strMessage"
    arrayModuloForm(134) = "    GoTo HandleExit"
    arrayModuloForm(135) = "End Function"
    arrayModuloForm(136) = ""
    arrayModuloForm(137) = "Public Sub FillList(ControlName As String, Values As Variant)"
    arrayModuloForm(138) = "    With Me.Controls(ControlName)"
    arrayModuloForm(139) = "        Dim iArrayForNext As Long"
    arrayModuloForm(140) = "        .Clear"
    arrayModuloForm(141) = "        For iArrayForNext = LBound(Values) To UBound(Values)"
    arrayModuloForm(142) = "            .AddItem Values(iArrayForNext)"
    arrayModuloForm(143) = "        Next"
    arrayModuloForm(144) = "    End With"
    arrayModuloForm(145) = "End Sub"
    arrayModuloForm(146) = ""
    arrayModuloForm(147) = "Private Function IsCorrectType(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(148) = "Dim strControlDataType As String, strMessage As String"
    arrayModuloForm(149) = "Dim dummy As Variant"
    arrayModuloForm(150) = "    strControlDataType = ControlDataType(ctl)"
    arrayModuloForm(151) = "On Error GoTo HandleError"
    arrayModuloForm(152) = "    Select Case strControlDataType"
    arrayModuloForm(153) = "    Case ""Boolean"""
    arrayModuloForm(154) = "        dummy = CBool(GetValue(ctl, strControlDataType))"
    arrayModuloForm(155) = "    Case ""Byte"""""
    arrayModuloForm(156) = "        dummy = CByte(GetValue(ctl, strControlDataType))"
    arrayModuloForm(157) = "    Case ""Currency"""
    arrayModuloForm(158) = "        dummy = CCur(GetValue(ctl, strControlDataType))"
    arrayModuloForm(159) = "    Case ""Date"""
    arrayModuloForm(160) = "        dummy = CDate(GetValue(ctl, strControlDataType))"
    arrayModuloForm(161) = "    Case ""Double"""
    arrayModuloForm(162) = "        dummy = CDbl(GetValue(ctl, strControlDataType))"
    arrayModuloForm(163) = "    Case ""Decimal"""
    arrayModuloForm(164) = "        dummy = CDec(GetValue(ctl, strControlDataType))"
    arrayModuloForm(165) = "    Case ""Integer"""
    arrayModuloForm(166) = "        dummy = CInt(GetValue(ctl, strControlDataType))"
    arrayModuloForm(167) = "    Case ""Long"""
    arrayModuloForm(168) = "        dummy = CLng(GetValue(ctl, strControlDataType))"
    arrayModuloForm(169) = "    Case ""Single"""
    arrayModuloForm(170) = "        dummy = CSng(GetValue(ctl, strControlDataType))"
    arrayModuloForm(171) = "    Case ""String"""
    arrayModuloForm(172) = "        dummy = CStr(GetValue(ctl, strControlDataType))"
    arrayModuloForm(173) = "    Case ""Variant"""
    arrayModuloForm(174) = "        dummy = CVar(GetValue(ctl, strControlDataType))"
    arrayModuloForm(175) = "    End Select"
    arrayModuloForm(176) = "    IsCorrectType = True"
    arrayModuloForm(177) = "HandleExit:"
    arrayModuloForm(178) = "    Exit Function"
    arrayModuloForm(179) = "HandleError:"
    arrayModuloForm(180) = "    IsCorrectType = False"
    arrayModuloForm(181) = "    Resume HandleExit"
    arrayModuloForm(182) = "End Function"
    arrayModuloForm(183) = ""
    arrayModuloForm(184) = "Private Function ControlName(ctl As MSForms.Control) As String"
    arrayModuloForm(185) = "On Error GoTo HandleError"
    arrayModuloForm(186) = "    If Not ctl Is Nothing Then"
    arrayModuloForm(187) = "        ControlName = ctl.Name"
    arrayModuloForm(188) = "        Select Case TypeName(ctl)"
    arrayModuloForm(189) = "        Case ""TextBox"", ""ListBox"", ""ComboBox"""
    arrayModuloForm(190) = "            If ctl.TabIndex > 0 Then"
    arrayModuloForm(191) = "                Dim c As MSForms.Control"
    arrayModuloForm(192) = "                For Each c In Me.Controls"
    arrayModuloForm(193) = "                    If c.TabIndex = ctl.TabIndex - 1 Then"
    arrayModuloForm(194) = "                        If TypeOf c Is MSForms.Label Then"
    arrayModuloForm(195) = "                            ControlName = c.Caption"
    arrayModuloForm(196) = "                        End If"
    arrayModuloForm(197) = "                    End If"
    arrayModuloForm(198) = "                Next"
    arrayModuloForm(199) = "            End If"
    arrayModuloForm(200) = "        Case Else"
    arrayModuloForm(201) = "            ControlName = ctl.Caption"
    arrayModuloForm(202) = "        End Select"
    arrayModuloForm(203) = "    End If"
    arrayModuloForm(204) = "HandleExit:"
    arrayModuloForm(205) = "    Exit Function"
    arrayModuloForm(206) = "HandleError:"
    arrayModuloForm(207) = "    Resume HandleExit"
    arrayModuloForm(208) = "End Function"
    arrayModuloForm(209) = ""
    arrayModuloForm(210) = "Private Function IsRequired(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(211) = "    Select Case ctl.Name"
    arrayModuloForm(212) = "    Case [CONTROLES_REQUERIDOS]"
    arrayModuloForm(213) = "        IsRequired = True"
    arrayModuloForm(214) = "    Case Else"
    arrayModuloForm(215) = "        IsRequired = False"
    arrayModuloForm(216) = "    End Select"
    arrayModuloForm(217) = "End Function"
    arrayModuloForm(218) = ""
    arrayModuloForm(219) = "Private Function IsInputControl(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(220) = "    Select Case TypeName(ctl)"
    arrayModuloForm(221) = "    Case ""TextBox"", ""ComboBox"", ""ListBox"", ""CheckBox"", ""OptionButton"", ""ToggleButton"""
    arrayModuloForm(222) = "        IsInputControl = True"
    arrayModuloForm(223) = "    Case Else"
    arrayModuloForm(224) = "        IsInputControl = False"
    arrayModuloForm(225) = "    End Select"
    arrayModuloForm(226) = "End Function"
    arrayModuloForm(227) = ""
    arrayModuloForm(228) = "Private Function HasValue(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(229) = "    Dim var As Variant"
    arrayModuloForm(230) = "    var = GetValue(ctl, ""Variant"")"
    arrayModuloForm(231) = "    If IsNull(var) Then"
    arrayModuloForm(232) = "        HasValue = False"
    arrayModuloForm(233) = "    ElseIf Len(var) = 0 Then"
    arrayModuloForm(234) = "        HasValue = False"
    arrayModuloForm(235) = "    Else"
    arrayModuloForm(236) = "        HasValue = True"
    arrayModuloForm(237) = "    End If"
    arrayModuloForm(238) = "End Function"
    arrayModuloForm(239) = ""
    arrayModuloForm(240) = "Private Function GetValue(ctl As MSForms.Control, strTypeName As String) As Variant"
    arrayModuloForm(241) = "On Error GoTo HandleError"
    arrayModuloForm(242) = "    Dim Value As Variant"
    arrayModuloForm(243) = "    Value = ctl.Value"
    arrayModuloForm(244) = "    If IsNull(Value) And strTypeName <> ""Variant"" Then"
    arrayModuloForm(245) = "        Select Case strTypeName"
    arrayModuloForm(246) = "        Case ""String"""
    arrayModuloForm(247) = "            Value = """""
    arrayModuloForm(248) = "        Case Else"
    arrayModuloForm(249) = "            Value = 0"
    arrayModuloForm(250) = "        End Select"
    arrayModuloForm(251) = "    End If"
    arrayModuloForm(252) = "HandleExit:"
    arrayModuloForm(253) = "    GetValue = Value"
    arrayModuloForm(254) = "    Exit Function"
    arrayModuloForm(255) = "HandleError:"
    arrayModuloForm(256) = "    Resume HandleExit"
    arrayModuloForm(257) = "End Function"
    arrayModuloForm(258) = ""
    arrayModuloForm(259) = "Private Sub SetValue(ctl As MSForms.Control, Value As Variant)"
    arrayModuloForm(260) = "On Error GoTo HandleError"
    arrayModuloForm(261) = "    ctl.Value = Value"
    arrayModuloForm(262) = "HandleExit:"
    arrayModuloForm(263) = "    Exit Sub"
    arrayModuloForm(264) = "HandleError:"
    arrayModuloForm(265) = "    Resume HandleExit"
    arrayModuloForm(266) = "End Sub"
    
    arrayModuloFuncaoLimpaControles(1) = "Public Sub LimpaControles()"
    arrayModuloFuncaoLimpaControles(2) = "        SetValue Me.[NOME_CONTROLE], """""
    arrayModuloFuncaoLimpaControles(3) = "End Sub"
    
    arrayModuloFuncaoControlDataType(1) = "Private Function ControlDataType(ctl As MSForms.Control) As String"
    arrayModuloFuncaoControlDataType(2) = "    Select Case ctl.Name"
    arrayModuloFuncaoControlDataType(3) = "    'Case ""txtId"": ControlDataType = ""Integer"""
    arrayModuloFuncaoControlDataType(4) = "    Case ""[NOME_CONTROLE]"": ControlDataType = ""[TIPO_DADO_CONTROLE]"""
    arrayModuloFuncaoControlDataType(5) = "    End Select"
    arrayModuloFuncaoControlDataType(6) = "End Function"
    
    arrayModuloFuncaoSetValues(1) = "Public Sub SetValues(udt[NOME_FORM] As [NOME_FORM])"
    arrayModuloFuncaoSetValues(2) = "    With udt[NOME_FORM]"
    arrayModuloFuncaoSetValues(3) = "        SetValue Me.[NOME_CONTROLE], .[NOME_CAMPO]"
    arrayModuloFuncaoSetValues(4) = "    End With"
    arrayModuloFuncaoSetValues(5) = "    "
    arrayModuloFuncaoSetValues(6) = "    Set cls[NOME_FORM] = udt[NOME_FORM]"
    arrayModuloFuncaoSetValues(7) = "End Sub"
    
    arrayModuloFuncaoGetValues(1) = "Public Sub GetValues(ByRef udt[NOME_FORM] As [NOME_FORM])"
    arrayModuloFuncaoGetValues(2) = "    With udt[NOME_FORM]"
    arrayModuloFuncaoGetValues(3) = "        '.Id = GetValue(Me.txtId, TypeName(.Id))"
    arrayModuloFuncaoGetValues(4) = "        .[NOME_CAMPO] = GetValue(Me.[NOME_CONTROLE], TypeName(.[NOME_CAMPO]))"
    arrayModuloFuncaoGetValues(5) = "    End With"
    arrayModuloFuncaoGetValues(6) = "End Sub"
End Sub

Private Sub cboControle_Change()
    If lstColunas.ListIndex > 0 Then
        Linha = lstColunas.ListIndex
        lstColunas.List(Linha, 1) = cboControle.text
    End If
End Sub

Private Sub cbxRequerido_Click()
    If lstColunas.ListIndex > 0 Then
        Linha = lstColunas.ListIndex
        lstColunas.List(Linha, 2) = IIf(cbxRequerido.Value, "Sim", "Não")
    End If
End Sub

Private Sub lstColunas_Click()
    Dim Linha As Long
    If lstColunas.ListIndex > 0 Then
        Linha = lstColunas.ListIndex
        txtNome.text = lstColunas.List(Linha, 0)
        cboControle.text = lstColunas.List(Linha, 1)
        cbxRequerido.Value = IIf(lstColunas.List(Linha, 2) = "Sim", True, False)
        
        Dim eChave As Boolean
        eChave = (lstColunas.List(Linha, 3) = "Sim")
        cboControle.Enabled = Not eChave
        cbxRequerido.Enabled = Not eChave
    End If
End Sub

Private Sub UserForm_Initialize()
    cboControle.AddItem "TextBox"
    cboControle.AddItem "ComboBox"
    cboControle.AddItem "CheckBox"
    cboControle.AddItem "OptionButtion"
    Call Init
End Sub

Private Sub cmdGerarFormulario_Click()
    If Trim(txtNomeFormulario.text) <> "" Then
        If IsVarArrayEmpty(controles) Then
            MsgBox "E onde estão os campos?"
        Else
            Call CriarForm(Trim(txtNomeFormulario.text))
        End If
    Else
        MsgBox "O nome do formulário é requerido"
    End If
End Sub

Private Sub CriarForm(ByVal NomeForm As String)
     
    Dim MyUserForm As VBComponent
    Dim btnOk As MSForms.CommandButton
    Dim btnCancelar As MSForms.CommandButton
    Dim MyComboBox As MSForms.ComboBox
    Dim N As Integer, MaxWidth As Long
    
    NomeForm = "ufm" & NomeForm
     
    'verifica se o formulário exite
    For N = 1 To ActiveWorkbook.VBProject.VBComponents.Count
        If ActiveWorkbook.VBProject.VBComponents(N).name = NomeForm Then
            MsgBox "Já existe um formulário com o mesmo nome"
            Exit Sub
        End If
    Next N
     
     'Cria o userform
    Set MyUserForm = ActiveWorkbook.VBProject _
    .VBComponents.Add(vbext_ct_MSForm)
    With MyUserForm
        .Properties("Height") = lstColunas.ListCount * 45
        .Properties("Width") = 300
        On Error Resume Next
        .name = NomeForm
        .Properties("Caption") = "Formulário - " & NomeForm
    End With
    
    'cria os controles referentes aos campos
    Dim i As Long
    Dim j As Integer
    Dim margemTopo As Integer, margeTopoInicial, distanciaEntre As Integer, margemEsquerda As Integer, alturaControle As Integer, larguraControle As Integer
    margeTopoInicial = 10
    margemTopo = 10
    distanciaEntre = 2
    margemEsquerda = 10
    alturaControle = 18
    larguraControle = 200
    
    For i = 2 To UBound(controles)
        'Rótulo
        Set Label = MyUserForm.Designer.Controls.Add("Forms.Label.1")
        With Label
            .Caption = controles(i, colunaCampo)
            .name = "lbl" & controles(i, colunaCampo)
            .Left = margemEsquerda
            .Top = margemTopo
            .Height = alturaControle
            .Width = larguraControle
        End With
        'Controle
        Set TextBox = MyUserForm.Designer.Controls.Add("Forms.TextBox.1")
        With TextBox
            .name = "txt" & controles(i, colunaCampo)
            .Left = margemEsquerda
            .Top = margemTopo + alturaControle + distanciaEntre
            .Height = alturaControle
            .Width = larguraControle
        End With
        
        margemTopo = margemTopo + margeTopoInicial + (alturaControle * 2)
    Next i
    
'    '//Add a combo box on the form
'    Set MyComboBox = MyUserForm.Designer.Controls.Add("Forms.ComboBox.1")
'    With MyComboBox
'        .Name = "Combo1"
'        .Left = 10
'        .Top = 10
'        .Height = 16
'        .Width = 100
'    End With
     
     '//Add a Cancel button to the form
    Set btnCancelar = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With btnCancelar
        .Caption = "Cancelar"
        .Height = 24
        .Width = 72
        .Left = 215
        .Top = 10
    End With
     
     '//Add an OK button to the form
    Set btnOk = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With btnOk
        .Caption = "OK"
        .Height = 24
        .Width = 72
        .Left = 215
        .Top = 44
    End With
     
    'código do form
    With MyUserForm.CodeModule
        countOfLines = .countOfLines
        For i = 1 To UBound(arrayModuloForm)
            Call InsertLine(MyUserForm, ReplaceToken(arrayModuloForm(i)))
        Next i
        
        Dim nomeControle As String, tipoDadoControle As String, nomeCampo As String, linhaAInserir As String, linhaNomeControle As Long
    
        'função LimpaControles
        i = 1
        While i <= UBound(arrayModuloFuncaoLimpaControles)
            If InStr(1, arrayModuloFuncaoLimpaControles(i), "[NOME_CONTROLE]") > 0 Then
                'guarda a referencia da linha com o conteudo
                linhaNomeControle = i
                For j = 2 To UBound(controles)
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), controles(j, colunaControle))
                    linhaAInserir = Replace(arrayModuloFuncaoLimpaControles(linhaNomeControle), "[NOME_CONTROLE]", nomeControle)
                    Call InsertLine(MyUserForm, linhaAInserir)
                Next j
            Else
                Call InsertLine(MyUserForm, arrayModuloFuncaoLimpaControles(i))
            End If
            i = i + 1
        Wend
        
        'função ControlDataType
        i = 1
        While i <= UBound(arrayModuloFuncaoControlDataType)
            If InStr(1, arrayModuloFuncaoControlDataType(i), "[NOME_CONTROLE]") > 0 Then
                'guarda a referencia da linha com o conteudo
                linhaNomeControle = i
                For j = 2 To UBound(controles)
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j, colunaControle))
                    tipoDadoControle = ObtemTipoDadoCampo(controles(j, colunaControle))
                    linhaAInserir = Replace(arrayModuloFuncaoControlDataType(linhaNomeControle), "[NOME_CONTROLE]", nomeControle)
                    linhaAInserir = Replace(linhaAInserir, "[TIPO_DADO_CONTROLE]", tipoDadoControle)
                    Call InsertLine(MyUserForm, linhaAInserir)
                Next j
            Else
                Call InsertLine(MyUserForm, arrayModuloFuncaoControlDataType(i))
            End If
            i = i + 1
        Wend
        
        'função SetValues
        i = 1
        While i <= UBound(arrayModuloFuncaoSetValues)
            If InStr(1, arrayModuloFuncaoSetValues(i), "[NOME_CONTROLE]") > 0 Then
                'guarda a referencia da linha com o conteudo
                linhaNomeControle = i
                For j = 2 To UBound(controles)
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j, colunaControle))
                    nomeCampo = ObtemTipoDadoCampo(controles(j, colunaCampo))
                    linhaAInserir = Replace(arrayModuloFuncaoSetValues(linhaNomeControle), "[NOME_CONTROLE]", nomeControle)
                    linhaAInserir = Replace(linhaAInserir, "[NOME_CAMPO]", nomeCampo)
                    Call InsertLine(MyUserForm, linhaAInserir)
                Next j
            Else
                Call InsertLine(MyUserForm, ReplaceToken(arrayModuloFuncaoSetValues(i)))
            End If
            i = i + 1
        Wend
        
        'função GetValues
        i = 1
        While i <= UBound(arrayModuloFuncaoGetValues)
            If InStr(1, arrayModuloFuncaoGetValues(i), "[NOME_CONTROLE]") > 0 Then
                'guarda a referencia da linha com o conteudo
                linhaNomeControle = i
                For j = 2 To UBound(controles)
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j, colunaControle))
                    nomeCampo = ObtemTipoDadoCampo(controles(j, colunaCampo))
                    linhaAInserir = Replace(arrayModuloFuncaoGetValues(linhaNomeControle), "[NOME_CONTROLE]", nomeControle)
                    linhaAInserir = Replace(linhaAInserir, "[NOME_CAMPO]", nomeCampo)
                    Call InsertLine(MyUserForm, linhaAInserir)
                Next j
            Else
                Call InsertLine(MyUserForm, ReplaceToken(arrayModuloFuncaoGetValues(i)))
            End If
            i = i + 1
        Wend
    End With
    
    Debug.Print "CountOfLines :" & countOfLines
    
    MsgBox NomeForm & " gerado com sucesso"
    Unload Me
End Sub

Private Sub InsertLine(ByRef Form As VBComponent, ByVal Linha As String)
    countOfLines = countOfLines + 1
    Call Form.CodeModule.InsertLines(countOfLines, Linha)
    'Debug.Print Linha
End Sub

Private Function ReplaceToken(ByVal text As String)
    Dim i As Integer
    '[NOME_FORM]
    text = Replace(text, "[NOME_FORM]", Trim(txtNomeFormulario.text))
    '[CHAVE_PRIMARIA]
    text = Replace(text, "[CHAVE_PRIMARIA]", ChavePrimaria)
    '[CONTROLES_REQUERIDOS]
    Dim controlesRequeridos() As String, controlesRequeridosCount As Long, controlesRequeridosIndex As Long
    i = 1
    Do
        If controles(i, colunaRequerido) = "Sim" Then controlesRequeridosCount = controlesRequeridosCount + 1
        i = i + 1
    Loop While i <= UBound(controles)
        
    ReDim controlesRequeridos(1 To controlesRequeridosCount)
    controlesRequeridosIndex = 1
    For i = 2 To UBound(controles)
        If controles(i, colunaRequerido) = "Sim" Then
            controlesRequeridos(controlesRequeridosIndex) = """" & ObtemNomeControle(controles(i, colunaCampo), lstColunas.List(i - 1, colunaControle - 1)) & """"
            controlesRequeridosIndex = controlesRequeridosIndex + 1
        End If
    Next i
    text = Replace(text, "[CONTROLES_REQUERIDOS]", Join(controlesRequeridos, ","))
    
    ReplaceToken = text
End Function

Private Function ChavePrimaria() As String
    If nomeCampoChavePrimaria = "" Then
        Dim i As Integer
        For i = 2 To UBound(controles)
            If controles(i, colunaEchave) = "Sim" Then
                nomeCampoChavePrimaria = controles(i, colunaCampo)
                Exit For
            End If
        Next i
    End If
    
    ChavePrimaria = nomeCampoChavePrimaria
End Function

Private Function ObtemNomeControle(ByVal campo As String, ByVal controle As String) As String
    Dim prefixo As String
    Select Case controle
        Case "TextBox"
            prefixo = "txt"
        Case "ComboBox"
            prefixo = ""
        Case "ListBox"
            prefixo = "lst"
        Case "CheckBox"
            prefixo = "cbx"
        Case "OptionButton"
            prefixo = "opt"
        Case "ToggleButton"
            prefixo = "tgb"
        Case Else
            prexixo = "ctl"
    End Select
    
    ObtemNomeControle = prefixo & campo
End Function

Private Function ObtemTipoDadoCampo(ByVal tipo As String) As String
    Select Case tipo
        Case "Text"
            tipo = "String"
        Case "Memo"
            tipo = "String"
        Case "Number"
            tipo = "Long"
        Case "Date/Time"
            tipo = "Date"
        Case "Currency"
            tipo = "Double"
        Case "AutoNumber"
            tipo = "Long"
        Case "Yes/No"
            tipo = "Boolean"
        Case "Hyperlink"
            tipo = "String"
        Case Else
            tipo = "Variant"
    End Select
    ObtemTipoDadoCampo = tipo
End Function

