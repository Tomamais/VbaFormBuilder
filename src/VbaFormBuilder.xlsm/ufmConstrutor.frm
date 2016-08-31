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
    arrayModuloForm(19) = "    btnCancelar.Enabled = Edicao"
    arrayModuloForm(20) = "    btnPrimeiro.Enabled = Not Edicao"
    arrayModuloForm(21) = "    btnAnterior.Enabled = Not Edicao"
    arrayModuloForm(22) = "    btnProximo.Enabled = Not Edicao"
    arrayModuloForm(23) = "    btnUltimo.Enabled = Not Edicao"
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
    arrayModuloForm(56) = "Private Sub btnUltimo_Click()"
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
    arrayModuloForm(82) = "Private Sub btnCancelar_Click()"
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
    arrayModuloForm(155) = "    Case ""Byte"""
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

Private Sub CriarForm(ByVal NomeEntidade As String)
     
    Dim newBook As Workbook
    Set newBook = Application.Workbooks.Add
    Dim MyUserForm As VBComponent
    Dim NomeForm As String
    NomeForm = NomeEntidade
    
    'botões
    Dim btnOk As MSForms.CommandButton
    Dim btnCancelar As MSForms.CommandButton
    Dim btnPesquisar As MSForms.CommandButton
    Dim btnPrimeiro As MSForms.CommandButton
    Dim btnAnterior As MSForms.CommandButton
    Dim btnProximo As MSForms.CommandButton
    Dim btnUltimo As MSForms.CommandButton
    'options
    Dim optNovo As MSForms.OptionButton
    Dim optAlterar As MSForms.OptionButton
    Dim optExcluir As MSForms.OptionButton
    'labels
    Dim lblStatus As MSForms.Label
    Dim N As Integer, MaxWidth As Long
    Dim nomeControle As String
    Dim tipoDadoControle As String
    Dim tipoControle As String
    Dim nomeCampo As String
    Dim linhaAInserir As String
    Dim linhaNomeControle As Long
    Dim nomeCampoPrivado As String
    Dim i As Long
    Dim j As Integer
    Dim margemTopo As Integer
    Dim margeTopoInicial
    Dim distanciaEntre As Integer
    Dim margemEsquerda As Integer
    Dim alturaControle As Integer
    Dim larguraControle As Integer
    
    countOfLines = 0
    
    'gera a classe
    Dim modEntidade As VBComponent
    Set modEntidade = newBook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    modEntidade.name = "mod" & NomeForm
    
    Call InsertLine(modEntidade, "Sub AbreForm" & NomeForm & "()")
    Call InsertLine(modEntidade, "    'variável do tipo da Classe " & NomeForm)
    Call InsertLine(modEntidade, "    Dim udt" & NomeForm & " As " & NomeForm)
    Call InsertLine(modEntidade, "    'Cria a isntância")
    Call InsertLine(modEntidade, "    Set udt" & NomeForm & " = New " & NomeForm)
    Call InsertLine(modEntidade, "    ")
    Call InsertLine(modEntidade, "    udt" & NomeForm & ".MoveLast")
    Call InsertLine(modEntidade, "    udt" & NomeForm & ".MoveFirst")
    Call InsertLine(modEntidade, "    'Atribui uma instância da classe " & NomeForm & " ao form")
    Call InsertLine(modEntidade, "    ufm" & NomeForm & ".SetValues udt" & NomeForm)
    Call InsertLine(modEntidade, "    'Mostra o form")
    Call InsertLine(modEntidade, "    ufm" & NomeForm & ".Show")
    Call InsertLine(modEntidade, "End Sub")
    
    countOfLines = 0
    
    'gera a classe
    Dim modAuxiliar As VBComponent
    Set modAuxiliar = newBook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    modAuxiliar.name = "modAuxiliar"
    
    Call InsertLine(modAuxiliar, "Public Function Nz(ByVal Value As Variant) As Variant")
    Call InsertLine(modAuxiliar, "    If IsNull(Value) Then Value = 0")
    Call InsertLine(modAuxiliar, "    Nz = Value")
    Call InsertLine(modAuxiliar, "End Function")

    countOfLines = 0
    
    Dim modTypes As VBComponent
    Set modTypes = newBook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    modTypes.name = "modTypes"
    
    Call InsertLine(modTypes, "Public Type Atendimento")
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        Call InsertLine(modTypes, "    " & nomeCampo & " As " & tipoDadoControle)
    Next i
    Call InsertLine(modTypes, "End Type")
    
    countOfLines = 0
    
    'gera a classe
    Dim classe As VBComponent
    Set classe = newBook.VBProject.VBComponents.Add(vbext_ct_ClassModule)
    classe.name = NomeForm
    
    Call InsertLine(classe, "Private mrstRecordset As Recordset")
    Call InsertLine(classe, "Private mbooLoaded As Boolean")
    Call InsertLine(classe, "Private mdbCurrentDb As Database")
    
    'campos privados
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        nomeControle = ObtemNomeControle(nomeCampo, lstColunas.List(i - 1, colunaControle - 1))
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        Call InsertLine(classe, "")
        Call InsertLine(classe, "Private " & nomeCampoPrivado & " As " & tipoDadoControle)
    Next i
    
    'propriedades
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        nomeControle = ObtemNomeControle(nomeCampo, lstColunas.List(i - 1, colunaControle - 1))
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        
        Call InsertLine(classe, "")
        Call InsertLine(classe, "Public Property Get " & nomeCampo & "() As " & tipoDadoControle)
        Call InsertLine(classe, "    " & nomeCampo & " = " & nomeCampoPrivado)
        Call InsertLine(classe, "End Property")
        If nomeCampo <> ChavePrimaria Then
            Call InsertLine(classe, "")
            Call InsertLine(classe, "Public Property Let " & nomeCampo & "(rData As " & tipoDadoControle & ")")
            Call InsertLine(classe, "    " & nomeCampoPrivado & " = rData")
            Call InsertLine(classe, "End Property")
        End If
    Next i
    
    Call InsertLine(classe, "Private Property Get Recordset() As Recordset")
    Call InsertLine(classe, "    Set Recordset = mrstRecordset")
    Call InsertLine(classe, "End Property")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "Private Property Set Recordset(rData As Recordset)")
    Call InsertLine(classe, "    Set mrstRecordset = rData")
    Call InsertLine(classe, "End Property")
    
    'função Load
    Call InsertLine(classe, "Private Sub Load()")
    Call InsertLine(classe, "    With Recordset")
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        If nomeCampo = ChavePrimaria Then
            Call InsertLine(classe, "        " & nomeCampoPrivado & " = Nz(.Fields(""" & nomeCampo & """).Value)")
        Else
            Call InsertLine(classe, "        Me." & nomeCampo & " = Nz(.Fields(""" & nomeCampo & """).Value)")
        End If
    Next i
    Call InsertLine(classe, "    End With")
    Call InsertLine(classe, "    mbooLoaded = True")
    Call InsertLine(classe, "End Sub")
    Call InsertLine(classe, "")
    
    'função Update
    Call InsertLine(classe, "Public Function Update() As Boolean")
    Call InsertLine(classe, "On Error GoTo HandleMessage")
    Call InsertLine(classe, "    Dim updated As Boolean")
    Call InsertLine(classe, "    With Recordset")
    Call InsertLine(classe, "        If mbooLoaded = True Then")
    Call InsertLine(classe, "            .Edit")
    Call InsertLine(classe, "        Else")
    Call InsertLine(classe, "            .AddNew")
    Call InsertLine(classe, "        End If")
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        eRequerido = controles(i, colunaRequerido) = "Sim"
        If nomeCampo = ChavePrimaria Then
            Call InsertLine(classe, "        " & nomeCampoPrivado & " = Nz(.Fields(""" & nomeCampo & """).Value)")
        Else
            If eRequerido Then
                Call InsertLine(classe, "        .Fields(""" & nomeCampo & """).Value = NullIfEmptyString(Me." & nomeCampo & ")")
            Else
                Call InsertLine(classe, "        .Fields(""" & nomeCampo & """).Value = Me." & nomeCampo)
            End If
        End If
    Next i

    Call InsertLine(classe, "        .Update")
    Call InsertLine(classe, "    End With")
    Call InsertLine(classe, "    mbooLoaded = True")
    Call InsertLine(classe, "    updated = True")
    Call InsertLine(classe, "HandleExit:")
    Call InsertLine(classe, "    Update = updated")
    Call InsertLine(classe, "    Exit Function")
    Call InsertLine(classe, "HandleMessage:")
    Call InsertLine(classe, "    MsgBox Err.Description")
    Call InsertLine(classe, "    GoTo HandleExit")
    Call InsertLine(classe, "End Function")
    Call InsertLine(classe, "")
    
    Call InsertLine(classe, "Public Property Get CurrentDb() As Database")
    Call InsertLine(classe, "    If mdbCurrentDb Is Nothing Then")
    Call InsertLine(classe, "        Set mdbCurrentDb = DBEngine.OpenDatabase(ThisWorkBook.Worksheets(""Config"").Range(""PASTA"") & ThisWorkBook.Worksheets(""Config"").Range(""ARQUIVO""))")
    Call InsertLine(classe, "    End If")
    Call InsertLine(classe, "    Set CurrentDb = mdbCurrentDb")
    Call InsertLine(classe, "End Property")
    Call InsertLine(classe, "")
    
    'demais funções
    Call InsertLine(classe, "Public Sub AddNew()")
    Call InsertLine(classe, "    mbooLoaded = False")
    Call InsertLine(classe, "End Sub")
    Call InsertLine(classe, "Public Function FindFirst(Optional Criteria As Variant) As Boolean")
    Call InsertLine(classe, "    If IsMissing(Criteria) Then")
    Call InsertLine(classe, "        Recordset.MoveFirst")
    Call InsertLine(classe, "        FindFirst = Not Recordset.EOF")
    Call InsertLine(classe, "    Else")
    Call InsertLine(classe, "        Recordset.FindFirst Criteria")
    Call InsertLine(classe, "        FindFirst = Not Recordset.NoMatch")
    Call InsertLine(classe, "    End If")
    Call InsertLine(classe, "    If FindFirst Then Load")
    Call InsertLine(classe, "End Function")
    Call InsertLine(classe, "Public Function FindLast(Optional Criteria As Variant) As Boolean")
    Call InsertLine(classe, "    If IsMissing(Criteria) Then")
    Call InsertLine(classe, "        Recordset.MoveLast")
    Call InsertLine(classe, "        FindLast = Not Recordset.EOF")
    Call InsertLine(classe, "    Else")
    Call InsertLine(classe, "        Recordset.FindLast Criteria")
    Call InsertLine(classe, "        FindLast = Not Recordset.NoMatch")
    Call InsertLine(classe, "    End If")
    Call InsertLine(classe, "    If FindLast Then Load")
    Call InsertLine(classe, "End Function")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "Public Function MoveFirst() As Boolean")
    Call InsertLine(classe, "    Recordset.MoveFirst")
    Call InsertLine(classe, "    Load")
    Call InsertLine(classe, "End Function")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "'Ocorre quando a classe é instanciada")
    Call InsertLine(classe, "Private Sub Class_Initialize()")
    Call InsertLine(classe, "    Set Recordset = CurrentDb.OpenRecordset(""" & NomeForm & """, dbOpenDynaset)")
    Call InsertLine(classe, "End Sub")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "'Ocorre quando a classe é tirada da memória (Set = Nothing)")
    Call InsertLine(classe, "Private Sub Class_Terminate()")
    Call InsertLine(classe, "    Recordset.Close")
    Call InsertLine(classe, "    Set Recordset = Nothing")
    Call InsertLine(classe, "End Sub")
    Call InsertLine(classe, "Function NullIfEmptyString(str As String) As Variant")
    Call InsertLine(classe, "    Dim strTrimmed As String: strTrimmed = Trim(str)")
    Call InsertLine(classe, "    If Len(strTrimmed) = 0 Then")
    Call InsertLine(classe, "        NullIfEmptyString = Null")
    Call InsertLine(classe, "    Else")
    Call InsertLine(classe, "        NullIfEmptyString = strTrimmed")
    Call InsertLine(classe, "    End If")
    Call InsertLine(classe, "End Function")
    Call InsertLine(classe, "Public Function MoveLast() As Boolean")
    Call InsertLine(classe, "    Recordset.MoveLast")
    Call InsertLine(classe, "    Load")
    Call InsertLine(classe, "End Function")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "Public Function MoveNext() As Boolean")
    Call InsertLine(classe, "    Dim result As Boolean")
    Call InsertLine(classe, "    If Not Recordset.EOF And Not (Recordset.AbsolutePosition + 1) >= Recordset.RecordCount Then")
    Call InsertLine(classe, "        Recordset.MoveNext")
    Call InsertLine(classe, "        Load")
    Call InsertLine(classe, "        result = True")
    Call InsertLine(classe, "    End If")
    Call InsertLine(classe, "    MoveNext = result")
    Call InsertLine(classe, "End Function")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "Public Function MovePrevious() As Boolean")
    Call InsertLine(classe, "    Dim result As Boolean")
    Call InsertLine(classe, "    If Not Recordset.BOF And Recordset.AbsolutePosition > 0 Then")
    Call InsertLine(classe, "        Recordset.MovePrevious")
    Call InsertLine(classe, "        Load")
    Call InsertLine(classe, "        result = True")
    Call InsertLine(classe, "    End If")
    Call InsertLine(classe, "    MovePrevious = result")
    Call InsertLine(classe, "End Function")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "Public Sub Delete()")
    Call InsertLine(classe, "    Recordset.Delete")
    Call InsertLine(classe, "End Sub")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "Public Property Get RecordSetWithFilter() As Recordset")
    Call InsertLine(classe, "    Dim rstFiltered As Recordset")
    Call InsertLine(classe, "    Set rstFiltered = Recordset.OpenRecordset")
    Call InsertLine(classe, "    If rstFiltered.RecordCount > 0 Then")
    Call InsertLine(classe, "        rstFiltered.MoveLast")
    Call InsertLine(classe, "        rstFiltered.MoveFirst")
    Call InsertLine(classe, "    End If")
    Call InsertLine(classe, "    Set RecordSetWithFilter = rstFiltered")
    Call InsertLine(classe, "End Property")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "Public Property Get Filter() As String")
    Call InsertLine(classe, "    Filter = Recordset.Filter")
    Call InsertLine(classe, "End Property")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "Public Property Let Filter(rFilter As String)")
    Call InsertLine(classe, "    Recordset.Filter = rFilter")
    Call InsertLine(classe, "End Property")

    NomeForm = "ufm" & NomeForm
     
    'verifica se o formulário exite
    For N = 1 To newBook.VBProject.VBComponents.Count
        If newBook.VBProject.VBComponents(N).name = NomeForm Then
            MsgBox "Já existe um formulário com o mesmo nome"
            Exit Sub
        End If
    Next N
     
     'Cria o userform
    Set MyUserForm = newBook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    With MyUserForm
        .Properties("Height") = lstColunas.ListCount * 45
        .Properties("Width") = 333
        On Error Resume Next
        .name = NomeForm
        .Properties("Caption") = "Formulário - " & NomeForm
    End With
    
    'cria os controles referentes aos campos
    margeTopoInicial = 10
    margemTopo = 10
    distanciaEntre = 2
    margemEsquerda = 10
    alturaControle = 18
    larguraControle = 200
    
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        eRequerido = controles(i, colunaRequerido) = "Sim"
        nomeControle = ObtemNomeControle(nomeCampo, lstColunas.List(i - 1, colunaControle - 1))
        tipoControle = lstColunas.List(i - 1, colunaControle - 1)
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
        Set TextBox = MyUserForm.Designer.Controls.Add("Forms." & tipoControle & ".1")
        With TextBox
            .name = nomeControle
            .Left = margemEsquerda
            .Top = margemTopo + alturaControle + distanciaEntre
            .Height = alturaControle
            .Width = larguraControle
        End With
        
        margemTopo = margemTopo + margeTopoInicial + (alturaControle * 2)
    Next i
 
    Set btnCancelar = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With btnCancelar
        .name = "btnCancelar"
        .Caption = "Cancelar"
        .Height = 24
        .Width = 72
        .Left = 234
        .Top = 24
    End With
     
    Set btnOk = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With btnOk
        .name = "btnOk"
        .Caption = "OK"
        .Height = 24
        .Width = 72
        .Left = 234
        .Top = 54
    End With
    
    Set optNovo = MyUserForm.Designer.Controls.Add("forms.OptionButton.1")
    With optNovo
        .name = "optNovo"
        .Caption = "Novo"
        .Height = 18
        .Width = 70
        .Left = 234
        .Top = 84
        .GroupName = "Operacoes"
    End With
    
    Set optAlterar = MyUserForm.Designer.Controls.Add("forms.OptionButton.1")
    With optAlterar
        .name = "optAlterar"
        .Caption = "Alterar"
        .Height = 18
        .Width = 70
        .Left = 234
        .Top = 102
        .GroupName = "Operacoes"
    End With
    
    Set optExcluir = MyUserForm.Designer.Controls.Add("forms.OptionButton.1")
    With optExcluir
        .name = "optExcluir"
        .Caption = "Excluir"
        .Height = 18
        .Width = 70
        .Left = 234
        .Top = 120
        .GroupName = "Operacoes"
    End With
    
    Set lblStatus = MyUserForm.Designer.Controls.Add("forms.Label.1")
    With lblStatus
        .name = "lblStatus"
        .Caption = ""
        .Height = 24
        .Width = 72
        .Left = 234
        .Top = 144
    End With
    
    Set btnPesquisar = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With btnPesquisar
        .name = "btnPesquisar"
        .Caption = "Pesquisar"
        .Height = 24
        .Width = 72
        .Left = 234
        .Top = 168
    End With
    
    Set btnPrimeiro = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With btnPrimeiro
        .name = "btnPrimeiro"
        .Caption = "Primeiro"
        .Height = 24
        .Width = 72
        .Left = 234
        .Top = 198
    End With
    
    Set btnAnterior = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With btnAnterior
        .name = "btnAnterior"
        .Caption = "Anterior"
        .Height = 24
        .Width = 72
        .Left = 234
        .Top = 228
    End With
    
    Set btnProximo = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With btnProximo
        .name = "btnProximo"
        .Caption = "Próximo"
        .Height = 24
        .Width = 72
        .Left = 234
        .Top = 258
    End With
    
    Set btnUltimo = MyUserForm.Designer.Controls.Add("forms.CommandButton.1")
    With btnUltimo
        .name = "btnUltimo"
        .Caption = "Último"
        .Height = 24
        .Width = 72
        .Left = 234
        .Top = 288
    End With
    
    'código do form
    With MyUserForm.CodeModule
        countOfLines = .countOfLines
        For i = 1 To UBound(arrayModuloForm)
            Call InsertLine(MyUserForm, ReplaceToken(arrayModuloForm(i)))
        Next i
        
        'função LimpaControles
        i = 1
        While i <= UBound(arrayModuloFuncaoLimpaControles)
            If InStr(1, arrayModuloFuncaoLimpaControles(i), "[NOME_CONTROLE]") > 0 Then
                'guarda a referencia da linha com o conteudo
                linhaNomeControle = i
                For j = 2 To UBound(controles)
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j - 1, colunaControle - 1))
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
                    If controles(j, colunaCampo) <> ChavePrimaria Then
                        nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j - 1, colunaControle - 1))
                        tipoDadoControle = ObtemTipoDadoCampo(controles(j, colunaControle))
                        linhaAInserir = Replace(arrayModuloFuncaoControlDataType(linhaNomeControle), "[NOME_CONTROLE]", nomeControle)
                        linhaAInserir = Replace(linhaAInserir, "[TIPO_DADO_CONTROLE]", tipoDadoControle)
                        Call InsertLine(MyUserForm, linhaAInserir)
                    End If
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
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j - 1, colunaControle - 1))
                    nomeCampo = controles(j, colunaCampo)
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
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j - 1, colunaControle - 1))
                    nomeCampo = controles(j, colunaCampo)
                        If nomeCampo <> ChavePrimaria Then
                        linhaAInserir = Replace(arrayModuloFuncaoGetValues(linhaNomeControle), "[NOME_CONTROLE]", nomeControle)
                        linhaAInserir = Replace(linhaAInserir, "[NOME_CAMPO]", nomeCampo)
                        Call InsertLine(MyUserForm, linhaAInserir)
                    End If
                Next j
            Else
                Call InsertLine(MyUserForm, ReplaceToken(arrayModuloFuncaoGetValues(i)))
            End If
            i = i + 1
        Wend
    End With
    
    'Formulário de Pesquisa
    Dim NomeFormPesquisa As String
    Dim UserFormPesquisa As VBComponent
    NomeFormPesquisa = NomeForm & "Pesquisa"
     
    'verifica se o formulário exite
    For N = 1 To newBook.VBProject.VBComponents.Count
        If newBook.VBProject.VBComponents(N).name = NomeFormPesquisa Then
            MsgBox "Já existe um formulário com o mesmo nome"
            Exit Sub
        End If
    Next N
     
    Dim alturaForm As Long
    alturaForm = lstColunas.ListCount * 45
     'Cria o userform
    Set UserFormPesquisa = newBook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    With UserFormPesquisa
        .Properties("Height") = alturaForm + 40
        .Properties("Width") = 600
        On Error Resume Next
        .name = NomeFormPesquisa
        .Properties("Caption") = "Formulário - " & NomeFormPesquisa
    End With
    
    'cria os controles referentes aos campos
    margeTopoInicial = 10
    margemTopo = 10
    distanciaEntre = 2
    margemEsquerda = 10
    alturaControle = 18
    larguraControle = 132
    
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        eRequerido = controles(i, colunaRequerido) = "Sim"
        nomeControle = ObtemNomeControle(nomeCampo, lstColunas.List(i - 1, colunaControle - 1))
        tipoControle = lstColunas.List(i - 1, colunaControle - 1)
        
        If tipoControle = "CheckBox" Then
            'Checkbox de filtro
            Set CheckBox = UserFormPesquisa.Designer.Controls.Add("Forms.CheckBox.1")
            With CheckBox
                .Caption = controles(i, colunaCampo)
                .name = nomeControle
                .Left = margemEsquerda
                .Top = margemTopo
                .Height = alturaControle
                .Width = larguraControle / 2
            End With
            'Checkbox de controle de filtro
            Set CheckBoxFiltro = UserFormPesquisa.Designer.Controls.Add("Forms.CheckBox.1")
            With CheckBoxFiltro
                .Caption = "Filtra " & controles(i, colunaCampo)
                .name = nomeControle & "Filtrar"
                .Left = margemEsquerda + (larguraControle / 2) + 5
                .Top = margemTopo
                .Height = alturaControle
                .Width = larguraControle / 2
            End With
        Else
            'Rótulo
            Set Label = UserFormPesquisa.Designer.Controls.Add("Forms.Label.1")
            With Label
                .Caption = controles(i, colunaCampo)
                .name = "lbl" & controles(i, colunaCampo)
                .Left = margemEsquerda
                .Top = margemTopo
                .Height = alturaControle
                .Width = larguraControle
            End With
            'Controle
            Set TextBox = UserFormPesquisa.Designer.Controls.Add("Forms." & tipoControle & ".1")
            With TextBox
                .name = nomeControle
                .Left = margemEsquerda
                .Top = margemTopo + alturaControle + distanciaEntre
                .Height = alturaControle
                .Width = larguraControle
            End With
        End If
        
        margemTopo = margemTopo + margeTopoInicial + (alturaControle * 2)
    Next i
 
    Set btnCancelar = UserFormPesquisa.Designer.Controls.Add("forms.CommandButton.1")
    With btnCancelar
        .name = "btnCancelar"
        .Caption = "Cancelar"
        .Height = 24
        .Width = 72
        .Left = 84.6
        .Top = alturaForm - 25
    End With
     
    Set btnOk = UserFormPesquisa.Designer.Controls.Add("forms.CommandButton.1")
    With btnOk
        .name = "btnOk"
        .Caption = "OK"
        .Height = 24
        .Width = 72
        .Left = 12.6
        .Top = alturaForm - 25
    End With
    
    Set lstLista = UserFormPesquisa.Designer.Controls.Add("forms.ListBox.1")
    With lstLista
        .name = "lst" & NomeEntidade
        .Height = alturaForm - 50
        .Width = 432
        .Left = 149.4
        .Top = 11
    End With
    
    Call InsertLine(UserFormPesquisa, "Private cls" & NomeEntidade & " As " & NomeEntidade)
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub btnCancelar_Click()")
    Call InsertLine(UserFormPesquisa, "    Unload Me")
    Call InsertLine(UserFormPesquisa, "End Sub")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub btnOK_Click()")
    Call InsertLine(UserFormPesquisa, "    If IsInputOk Then")
    Call InsertLine(UserFormPesquisa, "        Call PreencheListBox")
    Call InsertLine(UserFormPesquisa, "    End If")
    Call InsertLine(UserFormPesquisa, "End Sub")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub lst" & NomeEntidade & "_DblClick(ByVal Cancel As MSForms.ReturnBoolean)")
    Call InsertLine(UserFormPesquisa, "    If lst" & NomeEntidade & ".ListIndex > 0 Then")
    Call InsertLine(UserFormPesquisa, "        Dim Id As Long")
    Call InsertLine(UserFormPesquisa, "        " & ChavePrimaria & " = CLng(lst" & NomeEntidade & ".List(lst" & NomeEntidade & ".ListIndex, 0))")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "        cls" & NomeEntidade & ".MoveFirst")
    Call InsertLine(UserFormPesquisa, "        Do")
    Call InsertLine(UserFormPesquisa, "            If cls" & NomeEntidade & ".Id = Id Then")
    Call InsertLine(UserFormPesquisa, "                ufm" & NomeEntidade & ".SetValues cls" & NomeEntidade & "")
    Call InsertLine(UserFormPesquisa, "                Unload Me")
    Call InsertLine(UserFormPesquisa, "                Exit Do")
    Call InsertLine(UserFormPesquisa, "            End If")
    Call InsertLine(UserFormPesquisa, "        Loop While cls" & NomeEntidade & ".MoveNext")
    Call InsertLine(UserFormPesquisa, "    End If")
    Call InsertLine(UserFormPesquisa, "End Sub")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub UserForm_Initialize()")
    Call InsertLine(UserFormPesquisa, "    Set cls" & NomeEntidade & " = New " & NomeEntidade & "")
    Call InsertLine(UserFormPesquisa, "    cls" & NomeEntidade & ".MoveLast")
    Call InsertLine(UserFormPesquisa, "    cls" & NomeEntidade & ".MoveFirst")
    Call InsertLine(UserFormPesquisa, "    Call PreencheListBox")
    Call InsertLine(UserFormPesquisa, "End Sub")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub PreencheListBox()")
    Call InsertLine(UserFormPesquisa, "    Dim rstFiltro As Recordset")
    Call InsertLine(UserFormPesquisa, "    Dim arrayItems()")
    Call InsertLine(UserFormPesquisa, "    Dim filtros As String")
    Call InsertLine(UserFormPesquisa, "    Dim linha As Long")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "    lst" & NomeEntidade & ".Clear")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "    'aplica os filtros")
    
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        eRequerido = controles(i, colunaRequerido) = "Sim"
        nomeControle = ObtemNomeControle(nomeCampo, lstColunas.List(i - 1, colunaControle - 1))
        tipoControle = lstColunas.List(i - 1, colunaControle - 1)
        
        If nomeCampo = ChavePrimaria Then
            Call InsertLine(UserFormPesquisa, "    If Trim(" & nomeControle & ".Text) <> """" Then")
            Call InsertLine(UserFormPesquisa, "        filtros = """ & nomeCampo & " = "" & Trim(" & nomeControle & ".Text)")
            Call InsertLine(UserFormPesquisa, "    End If")
            Call InsertLine(UserFormPesquisa, "")
        Else
            If tipoDadoControle = "String" Then
                Call InsertLine(UserFormPesquisa, "    If Trim(" & nomeControle & ".Text) <> """" Then")
                Call InsertLine(UserFormPesquisa, "        If filtros <> """" Then filtros = filtros & "" AND """)
                Call InsertLine(UserFormPesquisa, "        filtros = filtros & """ & nomeCampo & " LIKE '*"" & Trim(" & nomeControle & ".Text) & ""*'""")
                Call InsertLine(UserFormPesquisa, "    End If")
                Call InsertLine(UserFormPesquisa, "")
            ElseIf tipoDadoControle = "Date" Then
                Call InsertLine(UserFormPesquisa, "    If Trim(" & nomeControle & ".Text) <> """" Then")
                Call InsertLine(UserFormPesquisa, "        If filtros <> """" Then filtros = filtros & "" AND """)
                Call InsertLine(UserFormPesquisa, "        filtros = filtros & """ & nomeCampo & " = #"" & Trim(CDate(" & nomeControle & ".Text)) & ""#""")
                Call InsertLine(UserFormPesquisa, "    End If")
                Call InsertLine(UserFormPesquisa, "")
            ElseIf tipoDadoControle = "Boolean" Then
                If tipoControle = "CheckBox" Then Call InsertLine(UserFormPesquisa, "    If " & nomeControle & "Filtrar.Value Then")
                Call InsertLine(UserFormPesquisa, "        If " & nomeControle & ".Value <> """" Then")
                Call InsertLine(UserFormPesquisa, "            If filtros <> """" Then filtros = filtros & "" AND """)
                Call InsertLine(UserFormPesquisa, "            filtros = filtros & """ & nomeCampo & " = "" & IIf(" & nomeControle & ".Value, ""True"", ""False"")")
                Call InsertLine(UserFormPesquisa, "        End If")
                If tipoControle = "CheckBox" Then Call InsertLine(UserFormPesquisa, "    End If")
            Else
                Call InsertLine(UserFormPesquisa, "    If Trim(" & nomeControle & ".Text) <> """" Then")
                Call InsertLine(UserFormPesquisa, "        If filtros <> """" Then filtros = filtros & "" AND """)
                Call InsertLine(UserFormPesquisa, "        filtros = """ & nomeCampo & " = "" & Trim(" & nomeControle & ".Text)")
                Call InsertLine(UserFormPesquisa, "    End If")
                Call InsertLine(UserFormPesquisa, "")
            End If
        End If
    Next i
    
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "    cls" & NomeEntidade & ".Filter = filtros")
    Call InsertLine(UserFormPesquisa, "    Set rstFiltro = cls" & NomeEntidade & ".RecordSetWithFilter")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "    linha = 1")
    Call InsertLine(UserFormPesquisa, "    ReDim arrayItems(1 To rstFiltro.RecordCount + 1, 1 To rstFiltro.Fields.Count)")
    Call InsertLine(UserFormPesquisa, "    Me.lst" & NomeEntidade & ".ColumnCount = rstFiltro.Fields.Count")
    Call InsertLine(UserFormPesquisa, "    'colunas")
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        Call InsertLine(UserFormPesquisa, "    arrayItems(linha, " & i - 1 & ") = """ & nomeCampo & """")
    Next i
    
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "    linha = 2")
    Call InsertLine(UserFormPesquisa, "    'linhas")
    Call InsertLine(UserFormPesquisa, "    Do While Not rstFiltro.EOF")
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        Call InsertLine(UserFormPesquisa, "        arrayItems(linha, " & i - 1 & ") = rstFiltro!" & nomeCampo)
    Next i
    
    Call InsertLine(UserFormPesquisa, "        linha = linha + 1")
    Call InsertLine(UserFormPesquisa, "        rstFiltro.MoveNext")
    Call InsertLine(UserFormPesquisa, "    Loop")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "    Me.lst" & NomeEntidade & ".List = arrayItems()")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "    cls" & NomeEntidade & ".Filter = """)
    Call InsertLine(UserFormPesquisa, "    Set rstFiltro = Nothing")
    Call InsertLine(UserFormPesquisa, "End Sub")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Function IsInputOk() As Boolean")
    Call InsertLine(UserFormPesquisa, "Dim ctl As MSForms.Control")
    Call InsertLine(UserFormPesquisa, "Dim strMessage As String")
    Call InsertLine(UserFormPesquisa, "    IsInputOk = False")
    Call InsertLine(UserFormPesquisa, "    For Each ctl In Me.Controls")
    Call InsertLine(UserFormPesquisa, "        If IsInputControl(ctl) Then")
    Call InsertLine(UserFormPesquisa, "            If IsRequired(ctl) Then")
    Call InsertLine(UserFormPesquisa, "                If Not HasValue(ctl) Then")
    Call InsertLine(UserFormPesquisa, "                    strMessage = ControlName(ctl) & "" é obrigatório""")
    Call InsertLine(UserFormPesquisa, "                End If")
    Call InsertLine(UserFormPesquisa, "            End If")
    Call InsertLine(UserFormPesquisa, "            If HasValue(ctl) Then")
    Call InsertLine(UserFormPesquisa, "                If Not IsCorrectType(ctl) Then")
    Call InsertLine(UserFormPesquisa, "                    strMessage = ControlName(ctl) & "" é inválido""")
    Call InsertLine(UserFormPesquisa, "                End If")
    Call InsertLine(UserFormPesquisa, "            End If")
    Call InsertLine(UserFormPesquisa, "        End If")
    Call InsertLine(UserFormPesquisa, "        If Len(strMessage) > 0 Then")
    Call InsertLine(UserFormPesquisa, "            ctl.SetFocus")
    Call InsertLine(UserFormPesquisa, "            GoTo HandleMessage")
    Call InsertLine(UserFormPesquisa, "        End If")
    Call InsertLine(UserFormPesquisa, "    Next")
    Call InsertLine(UserFormPesquisa, "    IsInputOk = True")
    Call InsertLine(UserFormPesquisa, "HandleExit:")
    Call InsertLine(UserFormPesquisa, "    Exit Function")
    Call InsertLine(UserFormPesquisa, "HandleMessage:")
    Call InsertLine(UserFormPesquisa, "    MsgBox strMessage")
    Call InsertLine(UserFormPesquisa, "    GoTo HandleExit")
    Call InsertLine(UserFormPesquisa, "End Function")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Public Sub FillList(ControlName As String, Values As Variant)")
    Call InsertLine(UserFormPesquisa, "    With Me.Controls(ControlName)")
    Call InsertLine(UserFormPesquisa, "        Dim iArrayForNext As Long")
    Call InsertLine(UserFormPesquisa, "        .Clear")
    Call InsertLine(UserFormPesquisa, "        For iArrayForNext = LBound(Values) To UBound(Values)")
    Call InsertLine(UserFormPesquisa, "            .AddItem Values(iArrayForNext)")
    Call InsertLine(UserFormPesquisa, "        Next")
    Call InsertLine(UserFormPesquisa, "    End With")
    Call InsertLine(UserFormPesquisa, "End Sub")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Function IsCorrectType(ctl As MSForms.Control) As Boolean")
    Call InsertLine(UserFormPesquisa, "Dim strControlDataType As String, strMessage As String")
    Call InsertLine(UserFormPesquisa, "Dim dummy As Variant")
    Call InsertLine(UserFormPesquisa, "    strControlDataType = ControlDataType(ctl)")
    Call InsertLine(UserFormPesquisa, "On Error GoTo HandleError")
    Call InsertLine(UserFormPesquisa, "    Select Case strControlDataType")
    Call InsertLine(UserFormPesquisa, "    Case ""Boolean""")
    Call InsertLine(UserFormPesquisa, "        dummy = CBool(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""Byte""")
    Call InsertLine(UserFormPesquisa, "        dummy = CByte(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""Currency""")
    Call InsertLine(UserFormPesquisa, "        dummy = CCur(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""Date""")
    Call InsertLine(UserFormPesquisa, "        dummy = CDate(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""Double""")
    Call InsertLine(UserFormPesquisa, "        dummy = CDbl(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""Decimal""")
    Call InsertLine(UserFormPesquisa, "        dummy = CDec(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""Integer""")
    Call InsertLine(UserFormPesquisa, "        dummy = CInt(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""Long""")
    Call InsertLine(UserFormPesquisa, "        dummy = CLng(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""Single""")
    Call InsertLine(UserFormPesquisa, "        dummy = CSng(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""String""")
    Call InsertLine(UserFormPesquisa, "        dummy = CStr(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    Case ""Variant""")
    Call InsertLine(UserFormPesquisa, "        dummy = CVar(GetValue(ctl, strControlDataType))")
    Call InsertLine(UserFormPesquisa, "    End Select")
    Call InsertLine(UserFormPesquisa, "    IsCorrectType = True")
    Call InsertLine(UserFormPesquisa, "HandleExit:")
    Call InsertLine(UserFormPesquisa, "    Exit Function")
    Call InsertLine(UserFormPesquisa, "HandleError:")
    Call InsertLine(UserFormPesquisa, "    IsCorrectType = False")
    Call InsertLine(UserFormPesquisa, "    Resume HandleExit")
    Call InsertLine(UserFormPesquisa, "End Function")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Function ControlDataType(ctl As MSForms.Control) As String")
    Call InsertLine(UserFormPesquisa, "    Select Case ctl.name")
    
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        eRequerido = controles(i, colunaRequerido) = "Sim"
        nomeControle = ObtemNomeControle(nomeCampo, lstColunas.List(i - 1, colunaControle - 1))
        tipoControle = lstColunas.List(i - 1, colunaControle - 1)
        
        Call InsertLine(UserFormPesquisa, "    Case """ & nomeControle & """: ControlDataType = """ & tipoDadoControle & """")
    Next i
    Call InsertLine(UserFormPesquisa, "    End Select")
    Call InsertLine(UserFormPesquisa, "End Function")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Function ControlName(ctl As MSForms.Control) As String")
    Call InsertLine(UserFormPesquisa, "On Error GoTo HandleError")
    Call InsertLine(UserFormPesquisa, "    If Not ctl Is Nothing Then")
    Call InsertLine(UserFormPesquisa, "        ControlName = ctl.name")
    Call InsertLine(UserFormPesquisa, "        Select Case TypeName(ctl)")
    Call InsertLine(UserFormPesquisa, "        Case ""TextBox"", ""ListBox"", ""ComboBox""")
    Call InsertLine(UserFormPesquisa, "            If ctl.TabIndex > 0 Then")
    Call InsertLine(UserFormPesquisa, "                Dim c As MSForms.Control")
    Call InsertLine(UserFormPesquisa, "                For Each c In Me.Controls")
    Call InsertLine(UserFormPesquisa, "                    If c.TabIndex = ctl.TabIndex - 1 Then")
    Call InsertLine(UserFormPesquisa, "                        If TypeOf c Is MSForms.Label Then")
    Call InsertLine(UserFormPesquisa, "                            ControlName = c.Caption")
    Call InsertLine(UserFormPesquisa, "                        End If")
    Call InsertLine(UserFormPesquisa, "                    End If")
    Call InsertLine(UserFormPesquisa, "                Next")
    Call InsertLine(UserFormPesquisa, "            End If")
    Call InsertLine(UserFormPesquisa, "        Case Else")
    Call InsertLine(UserFormPesquisa, "            ControlName = ctl.Caption")
    Call InsertLine(UserFormPesquisa, "        End Select")
    Call InsertLine(UserFormPesquisa, "    End If")
    Call InsertLine(UserFormPesquisa, "HandleExit:")
    Call InsertLine(UserFormPesquisa, "    Exit Function")
    Call InsertLine(UserFormPesquisa, "HandleError:")
    Call InsertLine(UserFormPesquisa, "    Resume HandleExit")
    Call InsertLine(UserFormPesquisa, "End Function")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Function IsRequired(ctl As MSForms.Control) As Boolean")
    Call InsertLine(UserFormPesquisa, "    Select Case ctl.name")
    Call InsertLine(UserFormPesquisa, ReplaceToken("    Case [CONTROLES_REQUERIDOS]"))
    Call InsertLine(UserFormPesquisa, "        IsRequired = True")
    Call InsertLine(UserFormPesquisa, "    Case Else")
    Call InsertLine(UserFormPesquisa, "        IsRequired = False")
    Call InsertLine(UserFormPesquisa, "    End Select")
    Call InsertLine(UserFormPesquisa, "End Function")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Function IsInputControl(ctl As MSForms.Control) As Boolean")
    Call InsertLine(UserFormPesquisa, "    Select Case TypeName(ctl)")
    Call InsertLine(UserFormPesquisa, "    Case ""TextBox"", ""ComboBox"", ""ListBox"", ""CheckBox"", ""OptionButton"", ""ToggleButton""")
    Call InsertLine(UserFormPesquisa, "        IsInputControl = True")
    Call InsertLine(UserFormPesquisa, "    Case Else")
    Call InsertLine(UserFormPesquisa, "        IsInputControl = False")
    Call InsertLine(UserFormPesquisa, "    End Select")
    Call InsertLine(UserFormPesquisa, "End Function")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Function HasValue(ctl As MSForms.Control) As Boolean")
    Call InsertLine(UserFormPesquisa, "    Dim var As Variant")
    Call InsertLine(UserFormPesquisa, "    var = GetValue(ctl, ""Variant"")")
    Call InsertLine(UserFormPesquisa, "    If IsNull(var) Then")
    Call InsertLine(UserFormPesquisa, "        HasValue = False")
    Call InsertLine(UserFormPesquisa, "    ElseIf Len(var) = 0 Then")
    Call InsertLine(UserFormPesquisa, "        HasValue = False")
    Call InsertLine(UserFormPesquisa, "    Else")
    Call InsertLine(UserFormPesquisa, "        HasValue = True")
    Call InsertLine(UserFormPesquisa, "    End If")
    Call InsertLine(UserFormPesquisa, "End Function")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Function GetValue(ctl As MSForms.Control, strTypeName As String) As Variant")
    Call InsertLine(UserFormPesquisa, "On Error GoTo HandleError")
    Call InsertLine(UserFormPesquisa, "    Dim Value As Variant")
    Call InsertLine(UserFormPesquisa, "    Value = ctl.Value")
    Call InsertLine(UserFormPesquisa, "    If IsNull(Value) And strTypeName <> ""Variant"" Then")
    Call InsertLine(UserFormPesquisa, "        Select Case strTypeName")
    Call InsertLine(UserFormPesquisa, "        Case ""String""")
    Call InsertLine(UserFormPesquisa, "            Value = """)
    Call InsertLine(UserFormPesquisa, "        Case Else")
    Call InsertLine(UserFormPesquisa, "            Value = 0")
    Call InsertLine(UserFormPesquisa, "        End Select")
    Call InsertLine(UserFormPesquisa, "    End If")
    Call InsertLine(UserFormPesquisa, "HandleExit:")
    Call InsertLine(UserFormPesquisa, "    GetValue = Value")
    Call InsertLine(UserFormPesquisa, "    Exit Function")
    Call InsertLine(UserFormPesquisa, "HandleError:")
    Call InsertLine(UserFormPesquisa, "    Resume HandleExit")
    Call InsertLine(UserFormPesquisa, "End Function")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub SetValue(ctl As MSForms.Control, Value As Variant)")
    Call InsertLine(UserFormPesquisa, "On Error GoTo HandleError")
    Call InsertLine(UserFormPesquisa, "    ctl.Value = Value")
    Call InsertLine(UserFormPesquisa, "HandleExit:")
    Call InsertLine(UserFormPesquisa, "    Exit Sub")
    Call InsertLine(UserFormPesquisa, "HandleError:")
    Call InsertLine(UserFormPesquisa, "    Resume HandleExit")
    Call InsertLine(UserFormPesquisa, "End Sub")
    
    
    'Adiciona as referencias no novo arquivo
    Dim ref As Reference
    For Each ref In ThisWorkbook.VBProject.References
        Call newBook.VBProject.References.AddFromGuid(ref.GUID, ref.Major, ref.Minor)
    Next ref
    
    'Define a planilha de configuração e valores
    Dim ws As Worksheet, rngPasta As Range, rngArquivo As Range
    Set ws = newBook.Worksheets(1)
    ws.name = "Config"
    Call newBook.Names.Add("PASTA", "=Config!R1C2")
    Call newBook.Names.Add("ARQUIVO", "=Config!R2C2")
    ws.Cells(1, 1).Value = "Pasta:"
    ws.Cells(2, 1).Value = "Arquivo:"
    Dim arrayArquivo() As String
    arrayArquivo = Split(ufmSelecionaBanco.txtCaminhoBanco.text, "\")
    ws.Cells(2, 2).Value = arrayArquivo(UBound(arrayArquivo))
    ws.Cells(1, 2).Value = Replace(ufmSelecionaBanco.txtCaminhoBanco.text, ws.Cells(2, 2).Value, "", Compare:=vbTextCompare)
    
    Debug.Print "CountOfLines :" & countOfLines
    
    MsgBox NomeForm & " gerado com sucesso"
    Unload Me
    Unload ufmSelecionaBanco
End Sub

Private Sub InsertLine(ByRef componente As VBComponent, ByVal Linha As String)
    countOfLines = countOfLines + 1
    Call componente.CodeModule.InsertLines(countOfLines, Linha)
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
            prefixo = "cbo"
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


Private Function ObtemAcronimoTipo(ByVal tipo As String) As String
    Select Case tipo
        Case "Integer"
            tipo = "int"
        Case "String"
            tipo = "str"
        Case "Long"
            tipo = "lng"
        Case "Date"
            tipo = "dt"
        Case "Double"
            tipo = "dbl"
        Case "Boolean"
            tipo = "boo"
        Case Else
            tipo = "var"
    End Select
    ObtemAcronimoTipo = tipo
End Function
