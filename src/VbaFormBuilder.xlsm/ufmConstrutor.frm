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
Dim controles()
Dim arrayModuloForm(1 To 319)

Public Sub DefineControles(ByRef pControles())
     controles = pControles
End Sub

Private Sub Init()
    arrayModuloForm(1) = "Public IsCancelled As Boolean"
    arrayModuloForm(2) = "Private clsAtendimento As Atendimento"
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
    arrayModuloForm(15) = "    txtId.Enabled = False"
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
    arrayModuloForm(28) = "    """
    arrayModuloForm(29) = "    If Not Edicao Then"
    arrayModuloForm(30) = "        optAlterar.Value = False"
    arrayModuloForm(31) = "        optExcluir.Value = False"
    arrayModuloForm(32) = "        optNovo.Value = False"
    arrayModuloForm(33) = "        lblStatus.Caption = """""""
    arrayModuloForm(34) = "    End If"
    arrayModuloForm(35) = "    "
    arrayModuloForm(36) = "    modoEdicao = Edicao"
    arrayModuloForm(37) = "End Sub"
    arrayModuloForm(38) = ""
    arrayModuloForm(39) = "Private Sub btnAnterior_Click()"
    arrayModuloForm(40) = "    If clsAtendimento.MovePrevious Then Call SetValues(clsAtendimento)"
    arrayModuloForm(41) = "End Sub"
    arrayModuloForm(42) = ""
    arrayModuloForm(43) = "Private Sub btnPesquisar_Click()"
    arrayModuloForm(44) = "    ufmAtendimentoPesquisa.Show"
    arrayModuloForm(45) = "End Sub"
    arrayModuloForm(46) = ""
    arrayModuloForm(47) = "Private Sub btnPrimeiro_Click()"
    arrayModuloForm(48) = "    clsAtendimento.MoveFirst"
    arrayModuloForm(49) = "    Call SetValues(clsAtendimento)"
    arrayModuloForm(50) = "End Sub"
    arrayModuloForm(51) = ""
    arrayModuloForm(52) = "Private Sub btnProximo_Click()"
    arrayModuloForm(53) = "    If clsAtendimento.MoveNext Then Call SetValues(clsAtendimento)"
    arrayModuloForm(54) = "End Sub"
    arrayModuloForm(55) = ""
    arrayModuloForm(56) = "Private Sub btnUtimo_Click()"
    arrayModuloForm(57) = "    clsAtendimento.MoveLast"
    arrayModuloForm(58) = "    Call SetValues(clsAtendimento)"
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
    arrayModuloForm(73) = "    clsAtendimento.AddNew"
    arrayModuloForm(74) = "End Sub"
    arrayModuloForm(75) = ""
    arrayModuloForm(76) = "Public Sub LimpaControles()"
    arrayModuloForm(77) = "        SetValue Me.txtId, """""""
    arrayModuloForm(78) = "        SetValue Me.txtCodigo, """""""
    arrayModuloForm(79) = "        SetValue Me.txtNome, """""""
    arrayModuloForm(80) = "        SetValue Me.txtAdmissao, """""""
    arrayModuloForm(81) = "        SetValue Me.txtNascimento, """""""
    arrayModuloForm(82) = "        SetValue Me.cbxPDD, """""""
    arrayModuloForm(83) = "        SetValue Me.txtTCasa, """""""
    arrayModuloForm(84) = "        SetValue Me.txtData, """""""
    arrayModuloForm(85) = "End Sub"
    arrayModuloForm(86) = ""
    arrayModuloForm(87) = "Private Sub UserForm_Initialize()"
    arrayModuloForm(88) = "    IsCancelled = True"
    arrayModuloForm(89) = "    "
    arrayModuloForm(90) = "    AlteraModo Edicao:=False"
    arrayModuloForm(91) = "End Sub"
    arrayModuloForm(92) = ""
    arrayModuloForm(93) = "Private Sub btnCancel_Click()"
    arrayModuloForm(94) = "    AlteraModo Edicao:=False"
    arrayModuloForm(95) = "    clsAtendimento.MovePrevious"
    arrayModuloForm(96) = "    Call SetValues(clsAtendimento)"
    arrayModuloForm(97) = "    'Me.Hide"
    arrayModuloForm(98) = "End Sub"
    arrayModuloForm(99) = ""
    arrayModuloForm(100) = "Private Sub btnOK_Click()"
    arrayModuloForm(101) = "    If optExcluir.Value Then"
    arrayModuloForm(102) = "        If MsgBox(""Deseja realmente excluir este registro?"", vbYesNo, ""Aviso de Exclusão"") = vbYes Then"
    arrayModuloForm(103) = "            clsAtendimento.Delete"
    arrayModuloForm(104) = "            AlteraModo Edicao:=False"
    arrayModuloForm(105) = "            clsAtendimento.MoveFirst"
    arrayModuloForm(106) = "            Call SetValues(clsAtendimento)"
    arrayModuloForm(107) = "        End If"
    arrayModuloForm(108) = "    ElseIf IsInputOk Then"
    arrayModuloForm(109) = "        IsCancelled = False"
    arrayModuloForm(110) = "        Call GetValues(clsAtendimento)"
    arrayModuloForm(111) = "        If clsAtendimento.Update Then"
    arrayModuloForm(112) = "            AlteraModo Edicao:=False"
    arrayModuloForm(113) = "            clsAtendimento.MoveFirst"
    arrayModuloForm(114) = "            Call SetValues(clsAtendimento)"
    arrayModuloForm(115) = "        End If"
    arrayModuloForm(116) = "        'Me.Hide"
    arrayModuloForm(117) = "    End If"
    arrayModuloForm(118) = "End Sub"
    arrayModuloForm(119) = ""
    arrayModuloForm(120) = "Public Sub SetValues(udtAtendimento As Atendimento)"
    arrayModuloForm(121) = "    With udtAtendimento"
    arrayModuloForm(122) = "        SetValue Me.txtId, .Id"
    arrayModuloForm(123) = "        SetValue Me.txtCodigo, .Codigo"
    arrayModuloForm(124) = "        SetValue Me.txtNome, .Nome"
    arrayModuloForm(125) = "        SetValue Me.txtAdmissao, .Admissao"
    arrayModuloForm(126) = "        SetValue Me.txtNascimento, .Nascimento"
    arrayModuloForm(127) = "        SetValue Me.cbxPDD, .PDD"
    arrayModuloForm(128) = "        SetValue Me.txtTCasa, .TCasa"
    arrayModuloForm(129) = "        SetValue Me.txtData, .Data"
    arrayModuloForm(130) = "    End With"
    arrayModuloForm(131) = "    "
    arrayModuloForm(132) = "    Set clsAtendimento = udtAtendimento"
    arrayModuloForm(133) = "End Sub"
    arrayModuloForm(134) = ""
    arrayModuloForm(135) = "Public Sub GetValues(ByRef udtAtendimento As Atendimento)"
    arrayModuloForm(136) = "    With udtAtendimento"
    arrayModuloForm(137) = "        '.Id = GetValue(Me.txtId, TypeName(.Id))"
    arrayModuloForm(138) = "        .Codigo = GetValue(Me.txtCodigo, TypeName(.Codigo))"
    arrayModuloForm(139) = "        .Nome = GetValue(Me.txtNome, TypeName(.Nome))"
    arrayModuloForm(140) = "        .Admissao = GetValue(Me.txtAdmissao, TypeName(.Admissao))"
    arrayModuloForm(141) = "        .Nascimento = GetValue(Me.txtNascimento, TypeName(.Nascimento))"
    arrayModuloForm(142) = "        .PDD = GetValue(Me.cbxPDD, TypeName(.PDD))"
    arrayModuloForm(143) = "        .TCasa = GetValue(Me.txtTCasa, TypeName(.TCasa))"
    arrayModuloForm(144) = "        .Data = GetValue(Me.txtData, TypeName(.Data))"
    arrayModuloForm(145) = "    End With"
    arrayModuloForm(146) = "End Sub"
    arrayModuloForm(147) = ""
    arrayModuloForm(148) = "Private Function IsInputOk() As Boolean"
    arrayModuloForm(149) = "Dim ctl As MSForms.Control"
    arrayModuloForm(150) = "Dim strMessage As String"
    arrayModuloForm(151) = "    IsInputOk = False"
    arrayModuloForm(152) = "    For Each ctl In Me.Controls"
    arrayModuloForm(153) = "        If IsInputControl(ctl) Then"
    arrayModuloForm(154) = "            If IsRequired(ctl) Then"
    arrayModuloForm(155) = "                If Not HasValue(ctl) Then"
    arrayModuloForm(156) = "                    strMessage = ControlName(ctl) & "" é obrigatório"""
    arrayModuloForm(157) = "                End If"
    arrayModuloForm(158) = "            End If"
    arrayModuloForm(159) = "            If Not IsCorrectType(ctl) Then"
    arrayModuloForm(160) = "                strMessage = ControlName(ctl) & "" é inválido"""
    arrayModuloForm(161) = "            End If"
    arrayModuloForm(162) = "        End If"
    arrayModuloForm(163) = "        If Len(strMessage) > 0 Then"
    arrayModuloForm(164) = "            ctl.SetFocus"
    arrayModuloForm(165) = "            GoTo HandleMessage"
    arrayModuloForm(166) = "        End If"
    arrayModuloForm(167) = "    Next"
    arrayModuloForm(168) = "    IsInputOk = True"
    arrayModuloForm(169) = "HandleExit:"
    arrayModuloForm(170) = "    Exit Function"
    arrayModuloForm(171) = "HandleMessage:"
    arrayModuloForm(172) = "    MsgBox strMessage"
    arrayModuloForm(173) = "    GoTo HandleExit"
    arrayModuloForm(174) = "End Function"
    arrayModuloForm(175) = ""
    arrayModuloForm(176) = "Public Sub FillList(ControlName As String, Values As Variant)"
    arrayModuloForm(177) = "    With Me.Controls(ControlName)"
    arrayModuloForm(178) = "        Dim iArrayForNext As Long"
    arrayModuloForm(179) = "        .Clear"
    arrayModuloForm(180) = "        For iArrayForNext = LBound(Values) To UBound(Values)"
    arrayModuloForm(181) = "            .AddItem Values(iArrayForNext)"
    arrayModuloForm(182) = "        Next"
    arrayModuloForm(183) = "    End With"
    arrayModuloForm(184) = "End Sub"
    arrayModuloForm(185) = ""
    arrayModuloForm(186) = "Private Function IsCorrectType(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(187) = "Dim strControlDataType As String, strMessage As String"
    arrayModuloForm(188) = "Dim dummy As Variant"
    arrayModuloForm(189) = "    strControlDataType = ControlDataType(ctl)"
    arrayModuloForm(190) = "On Error GoTo HandleError"
    arrayModuloForm(191) = "    Select Case strControlDataType"
    arrayModuloForm(192) = "    Case ""Boolean"""
    arrayModuloForm(193) = "        dummy = CBool(GetValue(ctl, strControlDataType))"
    arrayModuloForm(194) = "    Case ""Byte"""""
    arrayModuloForm(195) = "        dummy = CByte(GetValue(ctl, strControlDataType))"
    arrayModuloForm(196) = "    Case ""Currency"""
    arrayModuloForm(197) = "        dummy = CCur(GetValue(ctl, strControlDataType))"
    arrayModuloForm(198) = "    Case ""Date"""
    arrayModuloForm(199) = "        dummy = CDate(GetValue(ctl, strControlDataType))"
    arrayModuloForm(200) = "    Case ""Double"""
    arrayModuloForm(201) = "        dummy = CDbl(GetValue(ctl, strControlDataType))"
    arrayModuloForm(202) = "    Case ""Decimal"""
    arrayModuloForm(203) = "        dummy = CDec(GetValue(ctl, strControlDataType))"
    arrayModuloForm(204) = "    Case ""Integer"""
    arrayModuloForm(205) = "        dummy = CInt(GetValue(ctl, strControlDataType))"
    arrayModuloForm(206) = "    Case ""Long"""
    arrayModuloForm(207) = "        dummy = CLng(GetValue(ctl, strControlDataType))"
    arrayModuloForm(208) = "    Case ""Single"""
    arrayModuloForm(209) = "        dummy = CSng(GetValue(ctl, strControlDataType))"
    arrayModuloForm(210) = "    Case ""String"""
    arrayModuloForm(211) = "        dummy = CStr(GetValue(ctl, strControlDataType))"
    arrayModuloForm(212) = "    Case ""Variant"""
    arrayModuloForm(213) = "        dummy = CVar(GetValue(ctl, strControlDataType))"
    arrayModuloForm(214) = "    End Select"
    arrayModuloForm(215) = "    IsCorrectType = True"
    arrayModuloForm(216) = "HandleExit:"
    arrayModuloForm(217) = "    Exit Function"
    arrayModuloForm(218) = "HandleError:"
    arrayModuloForm(219) = "    IsCorrectType = False"
    arrayModuloForm(220) = "    Resume HandleExit"
    arrayModuloForm(221) = "End Function"
    arrayModuloForm(222) = ""
    arrayModuloForm(223) = "Private Function ControlDataType(ctl As MSForms.Control) As String"
    arrayModuloForm(224) = "    Select Case ctl.Name"
    arrayModuloForm(225) = "    'Case ""txtId"": ControlDataType = ""Integer"""
    arrayModuloForm(226) = "    Case ""txtCodigo"": ControlDataType = ""String"""
    arrayModuloForm(227) = "    Case ""txtNome""": ControlDataType = """String"""
    arrayModuloForm(228) = "    Case ""txtAdmissao""": ControlDataType = """Date"""
    arrayModuloForm(229) = "    Case ""txtNascimento""": ControlDataType = """Date"""
    arrayModuloForm(230) = "    Case ""cbxPDD"": ControlDataType = ""Boolean"""
    arrayModuloForm(231) = "    Case ""txtTCasa"": ControlDataType = ""Integer"""
    arrayModuloForm(232) = "    Case ""txtData"": ControlDataType = ""Date"""
    arrayModuloForm(233) = "    End Select"
    arrayModuloForm(234) = "End Function"
    arrayModuloForm(235) = ""
    arrayModuloForm(236) = "Private Function ControlName(ctl As MSForms.Control) As String"
    arrayModuloForm(237) = "On Error GoTo HandleError"
    arrayModuloForm(238) = "    If Not ctl Is Nothing Then"
    arrayModuloForm(239) = "        ControlName = ctl.Name"
    arrayModuloForm(240) = "        Select Case TypeName(ctl)"
    arrayModuloForm(241) = "        Case ""TextBox"", ""ListBox"", ""ComboBox"""
    arrayModuloForm(242) = "            If ctl.TabIndex > 0 Then"
    arrayModuloForm(243) = "                Dim c As MSForms.Control"
    arrayModuloForm(244) = "                For Each c In Me.Controls"
    arrayModuloForm(245) = "                    If c.TabIndex = ctl.TabIndex - 1 Then"
    arrayModuloForm(246) = "                        If TypeOf c Is MSForms.Label Then"
    arrayModuloForm(247) = "                            ControlName = c.Caption"
    arrayModuloForm(248) = "                        End If"
    arrayModuloForm(249) = "                    End If"
    arrayModuloForm(250) = "                Next"
    arrayModuloForm(251) = "            End If"
    arrayModuloForm(252) = "        Case Else"
    arrayModuloForm(253) = "            ControlName = ctl.Caption"
    arrayModuloForm(254) = "        End Select"
    arrayModuloForm(255) = "    End If"
    arrayModuloForm(256) = "HandleExit:"
    arrayModuloForm(257) = "    Exit Function"
    arrayModuloForm(258) = "HandleError:"
    arrayModuloForm(259) = "    Resume HandleExit"
    arrayModuloForm(260) = "End Function"
    arrayModuloForm(261) = ""
    arrayModuloForm(262) = "Private Function IsRequired(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(263) = "    Select Case ctl.Name"
    arrayModuloForm(264) = "    Case ""txtCodigo"", ""txtNome"", ""txtData"""
    arrayModuloForm(265) = "        IsRequired = True"
    arrayModuloForm(266) = "    Case Else"
    arrayModuloForm(267) = "        IsRequired = False"
    arrayModuloForm(268) = "    End Select"
    arrayModuloForm(269) = "End Function"
    arrayModuloForm(270) = ""
    arrayModuloForm(271) = "Private Function IsInputControl(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(272) = "    Select Case TypeName(ctl)"
    arrayModuloForm(273) = "    Case ""TextBox"", ""ComboBox"", ""ListBox"", ""CheckBox"", ""OptionButton"", ""ToggleButton"""
    arrayModuloForm(274) = "        IsInputControl = True"
    arrayModuloForm(275) = "    Case Else"
    arrayModuloForm(276) = "        IsInputControl = False"
    arrayModuloForm(277) = "    End Select"
    arrayModuloForm(278) = "End Function"
    arrayModuloForm(279) = ""
    arrayModuloForm(280) = "Private Function HasValue(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(281) = "    Dim var As Variant"
    arrayModuloForm(282) = "    var = GetValue(ctl, ""Variant"")"
    arrayModuloForm(283) = "    If IsNull(var) Then"
    arrayModuloForm(284) = "        HasValue = False"
    arrayModuloForm(285) = "    ElseIf Len(var) = 0 Then"
    arrayModuloForm(286) = "        HasValue = False"
    arrayModuloForm(287) = "    Else"
    arrayModuloForm(288) = "        HasValue = True"
    arrayModuloForm(289) = "    End If"
    arrayModuloForm(290) = "End Function"
    arrayModuloForm(291) = ""
    arrayModuloForm(292) = "Private Function GetValue(ctl As MSForms.Control, strTypeName As String) As Variant"
    arrayModuloForm(293) = "On Error GoTo HandleError"
    arrayModuloForm(294) = "    Dim Value As Variant"
    arrayModuloForm(295) = "    Value = ctl.Value"
    arrayModuloForm(296) = "    If IsNull(Value) And strTypeName <> ""Variant"" Then"
    arrayModuloForm(297) = "        Select Case strTypeName"
    arrayModuloForm(298) = "        Case ""String"""
    arrayModuloForm(299) = "            Value = """""
    arrayModuloForm(300) = "        Case Else"
    arrayModuloForm(301) = "            Value = 0"
    arrayModuloForm(302) = "        End Select"
    arrayModuloForm(303) = "    End If"
    arrayModuloForm(304) = "HandleExit:"
    arrayModuloForm(305) = "    GetValue = Value"
    arrayModuloForm(306) = "    Exit Function"
    arrayModuloForm(307) = "HandleError:"
    arrayModuloForm(308) = "    Resume HandleExit"
    arrayModuloForm(309) = "End Function"
    arrayModuloForm(310) = ""
    arrayModuloForm(311) = "Private Sub SetValue(ctl As MSForms.Control, Value As Variant)"
    arrayModuloForm(312) = "On Error GoTo HandleError"
    arrayModuloForm(313) = "    ctl.Value = Value"
    arrayModuloForm(314) = "HandleExit:"
    arrayModuloForm(315) = "    Exit Sub"
    arrayModuloForm(316) = "HandleError:"
    arrayModuloForm(317) = "    Resume HandleExit"
    arrayModuloForm(318) = "End Sub"
End Sub

Private Sub btnSelecionarRange_Click()
    Dim rangeSelecionado As Range
    Set rangeSelecionado = SelecionarRange()
    
    If Not rangeSelecionado Is Nothing Then
        lstColunas.Clear
        lstColunas.ColumnCount = 3
        ReDim controles(1 To rangeSelecionado.Columns.Count + 1, 1 To 3)
        Dim linha As Long
        linha = 1
        
        controles(linha, 1) = "Campo"
        controles(linha, 2) = "Controle"
        controles(linha, 3) = "Requerido"
        
        For linha = 1 To rangeSelecionado.Columns.Count
            Me.lstColunas.AddItem
            controles(linha + 1, 1) = rangeSelecionado.Cells(1, linha).Value
            controles(linha + 1, 2) = "TextBox"
            controles(linha + 1, 3) = "Não"
        Next linha
        
        lstColunas.List = controles
    End If
    
    Set rangeSelecionado = Nothing
End Sub

Private Sub cboControle_Change()
    If lstColunas.ListIndex > 0 Then
        linha = lstColunas.ListIndex
        lstColunas.List(linha, 1) = cboControle.Text
    End If
End Sub

Private Sub cbxRequerido_Click()
    If lstColunas.ListIndex > 0 Then
        linha = lstColunas.ListIndex
        lstColunas.List(linha, 2) = IIf(cbxRequerido.Value, "Sim", "Não")
    End If
End Sub

Private Sub lstColunas_Click()
    Dim linha As Long
    If lstColunas.ListIndex > 0 Then
        linha = lstColunas.ListIndex
        txtNome.Text = lstColunas.List(linha, 0)
        cboControle.Text = lstColunas.List(linha, 1)
        cbxRequerido.Value = IIf(lstColunas.List(linha, 2) = "Sim", True, False)
        
        Dim eChave As Boolean
        eChave = (lstColunas.List(linha, 3) = "Sim")
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
    If Trim(txtNomeFormulario.Text) <> "" Then
        If IsVarArrayEmpty(controles) Then
            MsgBox "E onde estão os campos?"
        Else
            Call CriarForm(Trim(txtNomeFormulario.Text))
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
    Dim N, X As Integer, MaxWidth As Long
    
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
            .Caption = controles(i, 1)
            .name = "lbl" & controles(i, 1)
            .Left = margemEsquerda
            .Top = margemTopo
            .Height = alturaControle
            .Width = larguraControle
        End With
        'Controle
        Set TextBox = MyUserForm.Designer.Controls.Add("Forms.TextBox.1")
        With TextBox
            .name = "txt" & controles(i, 1)
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
     
     '//Add code on the form for the CommandButtons
    With MyUserForm.codeModule
        X = .CountOfLines
        For i = 1 To UBound(arrayModuloForm)
            .InsertLines X + i, arrayModuloForm(i)
        Next i
    End With
    
    MsgBox NomeForm & " gerado com sucesso"
    Unload Me
End Sub
