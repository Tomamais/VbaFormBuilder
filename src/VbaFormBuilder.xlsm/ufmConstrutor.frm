VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmConstrutor 
   Caption         =   "Construtor de Formulários"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   11070
   OleObjectBlob   =   "ufmConstrutor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufmConstrutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tabelas()
Private controles()
Private fks()
Private arrayModuloForm(1 To 266)
Private arrayModuloFuncaoCleanControls(1 To 3)
Private arrayModuloFuncaoControlDataType(1 To 6)
Private arrayModuloFuncaoSetValues(1 To 7)
Private arrayModuloFuncaoGetValues(1 To 6)
Private arrayModuloFuncaoQueryClose(1 To 3)
Private arrayModuloFuncaoChangeMode(1 To 33)
Private nomeCampoChavePrimaria As String
Private countOfLines As Long
Private Const colunaCampo As Integer = 1
Private Const colunaControle As Integer = 2
Private Const colunaRequerido As Integer = 3
Private Const colunaEchave As Integer = 4
Private Const colunaRotulo As Integer = 5
Private Const colunaGerar As Integer = 6

Private Const colunaFKCampo As Integer = 1
Private Const colunaFKTabela As Integer = 2
Private Const colunaFKID As Integer = 3
Private Const colunaFKValor As Integer = 4
Private Const colunaEFK As Integer = 5

Public Sub DefineTabelas(ByRef pTabelas())
    tabelas = pTabelas
End Sub

Public Sub DefineControles(ByRef pControles())
    controles = pControles
    
    Dim i As Integer, j As Integer
    'cria uma fk para cada coluna
    
    Erase fks
    ReDim fks(1 To UBound(controles, 1), 1 To 5)
               
    fks(1, colunaFKCampo) = "Campo"
    fks(1, colunaFKTabela) = "Tabela"
    fks(1, colunaFKID) = "ID"
    fks(1, colunaFKValor) = "Valor"
    fks(1, colunaEFK) = "É FK"
    
    For j = 2 To UBound(controles, 1)
        fks(j, colunaFKCampo) = controles(j, colunaFKCampo)
        fks(j, colunaFKTabela) = ""              'controles(j, colunaFKTabela)
        fks(j, colunaFKID) = ""                  'controles(j, colunaFKID)
        fks(j, colunaFKValor) = ""               'controles(j, colunaFKValor)
        fks(j, colunaEFK) = "Não"
    Next j
    
End Sub

Private Sub Init()
    arrayModuloForm(1) = "Public IsCancelled As Boolean"
    arrayModuloForm(2) = "Private cls[NOME_ENTIDADE] As [NOME_ENTIDADE]"
    arrayModuloForm(3) = "Private modoEdicao As Boolean"
    arrayModuloForm(4) = ""
    arrayModuloForm(5) = "Private Sub btnAnterior_Click()"
    arrayModuloForm(6) = "    If cls[NOME_ENTIDADE].MovePrevious Then Call SetValues(cls[NOME_ENTIDADE])"
    arrayModuloForm(7) = "End Sub"
    arrayModuloForm(8) = ""
    arrayModuloForm(9) = "Private Sub btnPesquisar_Click()"
    arrayModuloForm(10) = "    ufm[NOME_FORM]Pesquisa.Show"
    arrayModuloForm(11) = "End Sub"
    arrayModuloForm(12) = ""
    arrayModuloForm(13) = "Private Sub btnPrimeiro_Click()"
    arrayModuloForm(14) = "    cls[NOME_ENTIDADE].MoveFirst"
    arrayModuloForm(15) = "    Call SetValues(cls[NOME_ENTIDADE])"
    arrayModuloForm(16) = "End Sub"
    arrayModuloForm(17) = ""
    arrayModuloForm(18) = "Private Sub btnProximo_Click()"
    arrayModuloForm(19) = "    If cls[NOME_ENTIDADE].MoveNext Then Call SetValues(cls[NOME_ENTIDADE])"
    arrayModuloForm(20) = "End Sub"
    arrayModuloForm(21) = ""
    arrayModuloForm(22) = "Private Sub btnUltimo_Click()"
    arrayModuloForm(23) = "    cls[NOME_ENTIDADE].MoveLast"
    arrayModuloForm(24) = "    Call SetValues(cls[NOME_ENTIDADE])"
    arrayModuloForm(25) = "End Sub"
    arrayModuloForm(26) = ""
    arrayModuloForm(27) = "Private Sub optAlterar_Click()"
    arrayModuloForm(28) = "    ChangeMode Edicao:=True"
    arrayModuloForm(29) = "End Sub"
    arrayModuloForm(30) = ""
    arrayModuloForm(31) = "Private Sub optExcluir_Click()"
    arrayModuloForm(32) = "    lblStatus.Caption = ""Modo de exclusão"""
    arrayModuloForm(33) = "    ChangeMode Edicao:=True"
    arrayModuloForm(34) = "End Sub"
    arrayModuloForm(35) = ""
    arrayModuloForm(36) = "Private Sub optNovo_Click()"
    arrayModuloForm(37) = "    ChangeMode Edicao:=True"
    arrayModuloForm(38) = "    Call CleanControls"
    arrayModuloForm(39) = "    cls[NOME_ENTIDADE].AddNew"
    arrayModuloForm(40) = "End Sub"
    arrayModuloForm(41) = ""
    arrayModuloForm(42) = "Private Sub UserForm_Initialize()"
    arrayModuloForm(43) = "    IsCancelled = True"
    arrayModuloForm(44) = "    Call LoadDependentCombos"
    arrayModuloForm(45) = "    ChangeMode Edicao:=False"
    arrayModuloForm(46) = "End Sub"
    arrayModuloForm(47) = ""
    arrayModuloForm(48) = "Private Sub btnCancelar_Click()"
    arrayModuloForm(49) = "    ChangeMode Edicao:=False"
    arrayModuloForm(50) = "    cls[NOME_ENTIDADE].MovePrevious"
    arrayModuloForm(51) = "    Call SetValues(cls[NOME_ENTIDADE])"
    arrayModuloForm(52) = "    'Me.Hide"
    arrayModuloForm(53) = "End Sub"
    arrayModuloForm(54) = ""
    arrayModuloForm(55) = "Private Sub btnOK_Click()"
    arrayModuloForm(56) = "    If optExcluir.Value Then"
    arrayModuloForm(57) = "        If MsgBox(""Deseja realmente excluir este registro?"", vbYesNo, ""Aviso de Exclusão"") = vbYes Then"
    arrayModuloForm(58) = "            cls[NOME_ENTIDADE].Delete"
    arrayModuloForm(59) = "            ChangeMode Edicao:=False"
    arrayModuloForm(60) = "            cls[NOME_ENTIDADE].MoveFirst"
    arrayModuloForm(61) = "            Call SetValues(cls[NOME_ENTIDADE])"
    arrayModuloForm(62) = "        End If"
    arrayModuloForm(63) = "    ElseIf IsInputOk Then"
    arrayModuloForm(64) = "        IsCancelled = False"
    arrayModuloForm(65) = "        Call GetValues(cls[NOME_ENTIDADE])"
    arrayModuloForm(66) = "        If cls[NOME_ENTIDADE].Update Then"
    arrayModuloForm(67) = "            ChangeMode Edicao:=False"
    arrayModuloForm(68) = "            cls[NOME_ENTIDADE].MoveFirst"
    arrayModuloForm(69) = "            Call SetValues(cls[NOME_ENTIDADE])"
    arrayModuloForm(70) = "        End If"
    arrayModuloForm(71) = "        'Me.Hide"
    arrayModuloForm(72) = "    End If"
    arrayModuloForm(73) = "End Sub"
    arrayModuloForm(74) = ""
    arrayModuloForm(75) = "Private Function IsInputOk() As Boolean"
    arrayModuloForm(76) = "Dim ctl As MSForms.Control"
    arrayModuloForm(77) = "Dim strMessage As String"
    arrayModuloForm(78) = "    IsInputOk = False"
    arrayModuloForm(79) = "    For Each ctl In Me.Controls"
    arrayModuloForm(80) = "        If IsInputControl(ctl) Then"
    arrayModuloForm(81) = "            If IsRequired(ctl) Then"
    arrayModuloForm(82) = "                If Not HasValue(ctl) Then"
    arrayModuloForm(83) = "                    strMessage = ControlName(ctl) & "" é obrigatório"""
    arrayModuloForm(84) = "                End If"
    arrayModuloForm(85) = "            End If"
    arrayModuloForm(86) = "            If Not IsCorrectType(ctl) Then"
    arrayModuloForm(87) = "                strMessage = ControlName(ctl) & "" é inválido"""
    arrayModuloForm(88) = "            End If"
    arrayModuloForm(89) = "        End If"
    arrayModuloForm(90) = "        If Len(strMessage) > 0 Then"
    arrayModuloForm(91) = "            ctl.SetFocus"
    arrayModuloForm(92) = "            GoTo HandleMessage"
    arrayModuloForm(93) = "        End If"
    arrayModuloForm(94) = "    Next"
    arrayModuloForm(95) = "    IsInputOk = True"
    arrayModuloForm(96) = "HandleExit:"
    arrayModuloForm(97) = "    Exit Function"
    arrayModuloForm(98) = "HandleMessage:"
    arrayModuloForm(99) = "    MsgBox strMessage"
    arrayModuloForm(100) = "    GoTo HandleExit"
    arrayModuloForm(101) = "End Function"
    arrayModuloForm(102) = ""
    arrayModuloForm(103) = "Public Sub FillList(ControlName As String, Values As Variant)"
    arrayModuloForm(104) = "    With Me.Controls(ControlName)"
    arrayModuloForm(105) = "        Dim iArrayForNext As Long"
    arrayModuloForm(106) = "        .Clear"
    arrayModuloForm(107) = "        For iArrayForNext = LBound(Values) To UBound(Values)"
    arrayModuloForm(108) = "            .AddItem Values(iArrayForNext)"
    arrayModuloForm(109) = "        Next"
    arrayModuloForm(110) = "    End With"
    arrayModuloForm(111) = "End Sub"
    arrayModuloForm(112) = ""
    arrayModuloForm(113) = "Private Function IsCorrectType(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(114) = "Dim strControlDataType As String, strMessage As String"
    arrayModuloForm(115) = "Dim dummy As Variant"
    arrayModuloForm(116) = "    strControlDataType = ControlDataType(ctl)"
    arrayModuloForm(117) = "On Error GoTo HandleError"
    arrayModuloForm(118) = "    Select Case strControlDataType"
    arrayModuloForm(119) = "    Case ""Boolean"""
    arrayModuloForm(120) = "        dummy = CBool(GetValue(ctl, strControlDataType))"
    arrayModuloForm(121) = "    Case ""Byte"""
    arrayModuloForm(122) = "        dummy = CByte(GetValue(ctl, strControlDataType))"
    arrayModuloForm(123) = "    Case ""Currency"""
    arrayModuloForm(124) = "        dummy = CCur(GetValue(ctl, strControlDataType))"
    arrayModuloForm(125) = "    Case ""Date"""
    arrayModuloForm(126) = "        dummy = CDate(GetValue(ctl, strControlDataType))"
    arrayModuloForm(127) = "    Case ""Double"""
    arrayModuloForm(128) = "        dummy = CDbl(GetValue(ctl, strControlDataType))"
    arrayModuloForm(129) = "    Case ""Decimal"""
    arrayModuloForm(130) = "        dummy = CDec(GetValue(ctl, strControlDataType))"
    arrayModuloForm(131) = "    Case ""Integer"""
    arrayModuloForm(132) = "        dummy = CInt(GetValue(ctl, strControlDataType))"
    arrayModuloForm(133) = "    Case ""Long"""
    arrayModuloForm(134) = "        dummy = CLng(GetValue(ctl, strControlDataType))"
    arrayModuloForm(135) = "    Case ""Single"""
    arrayModuloForm(136) = "        dummy = CSng(GetValue(ctl, strControlDataType))"
    arrayModuloForm(137) = "    Case ""String"""
    arrayModuloForm(138) = "        dummy = CStr(GetValue(ctl, strControlDataType))"
    arrayModuloForm(139) = "    Case ""Variant"""
    arrayModuloForm(140) = "        dummy = CVar(GetValue(ctl, strControlDataType))"
    arrayModuloForm(141) = "    End Select"
    arrayModuloForm(142) = "    IsCorrectType = True"
    arrayModuloForm(143) = "HandleExit:"
    arrayModuloForm(144) = "    Exit Function"
    arrayModuloForm(145) = "HandleError:"
    arrayModuloForm(146) = "    IsCorrectType = False"
    arrayModuloForm(147) = "    Resume HandleExit"
    arrayModuloForm(148) = "End Function"
    arrayModuloForm(149) = ""
    arrayModuloForm(150) = "Private Function ControlName(ctl As MSForms.Control) As String"
    arrayModuloForm(151) = "On Error GoTo HandleError"
    arrayModuloForm(152) = "    If Not ctl Is Nothing Then"
    arrayModuloForm(153) = "        ControlName = ctl.Name"
    arrayModuloForm(154) = "        Select Case TypeName(ctl)"
    arrayModuloForm(155) = "        Case ""TextBox"", ""ListBox"", ""ComboBox"""
    arrayModuloForm(156) = "            If ctl.TabIndex > 0 Then"
    arrayModuloForm(157) = "                Dim c As MSForms.Control"
    arrayModuloForm(158) = "                For Each c In Me.Controls"
    arrayModuloForm(159) = "                    If c.TabIndex = ctl.TabIndex - 1 Then"
    arrayModuloForm(160) = "                        If TypeOf c Is MSForms.Label Then"
    arrayModuloForm(161) = "                            ControlName = c.Caption"
    arrayModuloForm(162) = "                        End If"
    arrayModuloForm(163) = "                    End If"
    arrayModuloForm(164) = "                Next"
    arrayModuloForm(165) = "            End If"
    arrayModuloForm(166) = "        Case Else"
    arrayModuloForm(167) = "            ControlName = ctl.Caption"
    arrayModuloForm(168) = "        End Select"
    arrayModuloForm(169) = "    End If"
    arrayModuloForm(170) = "HandleExit:"
    arrayModuloForm(171) = "    Exit Function"
    arrayModuloForm(172) = "HandleError:"
    arrayModuloForm(173) = "    Resume HandleExit"
    arrayModuloForm(174) = "End Function"
    arrayModuloForm(175) = ""
    arrayModuloForm(176) = "Private Function IsRequired(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(177) = "    Select Case ctl.Name"
    arrayModuloForm(178) = "    Case [CONTROLES_REQUERIDOS]"
    arrayModuloForm(179) = "        IsRequired = True"
    arrayModuloForm(180) = "    Case Else"
    arrayModuloForm(181) = "        IsRequired = False"
    arrayModuloForm(182) = "    End Select"
    arrayModuloForm(183) = "End Function"
    arrayModuloForm(184) = ""
    arrayModuloForm(185) = "Private Function IsInputControl(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(186) = "    Select Case TypeName(ctl)"
    arrayModuloForm(187) = "    Case ""TextBox"", ""ComboBox"", ""ListBox"", ""CheckBox"", ""OptionButton"", ""ToggleButton"""
    arrayModuloForm(188) = "        IsInputControl = True"
    arrayModuloForm(189) = "    Case Else"
    arrayModuloForm(190) = "        IsInputControl = False"
    arrayModuloForm(191) = "    End Select"
    arrayModuloForm(192) = "End Function"
    arrayModuloForm(193) = ""
    arrayModuloForm(194) = "Private Function HasValue(ctl As MSForms.Control) As Boolean"
    arrayModuloForm(195) = "    Dim var As Variant"
    arrayModuloForm(196) = "    var = GetValue(ctl, ""Variant"")"
    arrayModuloForm(197) = "    If IsNull(var) Then"
    arrayModuloForm(198) = "        HasValue = False"
    arrayModuloForm(199) = "    ElseIf Len(var) = 0 Then"
    arrayModuloForm(200) = "        HasValue = False"
    arrayModuloForm(201) = "    Else"
    arrayModuloForm(202) = "        HasValue = True"
    arrayModuloForm(203) = "    End If"
    arrayModuloForm(204) = "End Function"
    arrayModuloForm(205) = ""
    arrayModuloForm(206) = "Private Function GetValue(ctl As MSForms.Control, strTypeName As String) As Variant"
    arrayModuloForm(207) = "On Error GoTo HandleError"
    arrayModuloForm(208) = "    Dim Value As Variant"
    arrayModuloForm(209) = "    Value = ctl.Value"
    arrayModuloForm(210) = "    If IsNull(Value) And strTypeName <> ""Variant"" Then"
    arrayModuloForm(211) = "        Select Case strTypeName"
    arrayModuloForm(212) = "        Case ""String"""
    arrayModuloForm(213) = "            Value = """""
    arrayModuloForm(214) = "        Case Else"
    arrayModuloForm(215) = "            Value = 0"
    arrayModuloForm(216) = "        End Select"
    arrayModuloForm(217) = "    End If"
    arrayModuloForm(218) = "HandleExit:"
    arrayModuloForm(219) = "    GetValue = Value"
    arrayModuloForm(220) = "    Exit Function"
    arrayModuloForm(221) = "HandleError:"
    arrayModuloForm(222) = "    Resume HandleExit"
    arrayModuloForm(223) = "End Function"
    arrayModuloForm(224) = ""
    arrayModuloForm(225) = "Private Sub SetValue(ctl As MSForms.Control, Value As Variant)"
    arrayModuloForm(226) = "On Error GoTo HandleError"
    arrayModuloForm(227) = "    ctl.Value = Value"
    arrayModuloForm(228) = "HandleExit:"
    arrayModuloForm(229) = "    Exit Sub"
    arrayModuloForm(230) = "HandleError:"
    arrayModuloForm(231) = "    Resume HandleExit"
    arrayModuloForm(232) = "End Sub"
    
    arrayModuloFuncaoChangeMode(1) = "Private Sub ChangeMode(ByVal Edicao As Boolean)"
    arrayModuloFuncaoChangeMode(2) = "    Dim ctl As MSForms.Control"
    arrayModuloFuncaoChangeMode(3) = "    'controles de input"
    arrayModuloFuncaoChangeMode(4) = "    For Each ctl In Me.Controls"
    arrayModuloFuncaoChangeMode(5) = "        If IsInputControl(ctl) Then"
    arrayModuloFuncaoChangeMode(6) = "           ctl.Enabled = Edicao"
    arrayModuloFuncaoChangeMode(7) = "        End If"
    arrayModuloFuncaoChangeMode(8) = "    Next"
    arrayModuloFuncaoChangeMode(9) = "    "
    arrayModuloFuncaoChangeMode(10) = "    'excessão"
    arrayModuloFuncaoChangeMode(11) = "    txt[CHAVE_PRIMARIA].Enabled = False"
    arrayModuloFuncaoChangeMode(12) = "    "
    arrayModuloFuncaoChangeMode(13) = "    'botoes de navegacao"
    arrayModuloFuncaoChangeMode(14) = "    btnOk.Enabled = Edicao"
    arrayModuloFuncaoChangeMode(15) = "    btnCancelar.Enabled = Edicao"
    arrayModuloFuncaoChangeMode(16) = "    btnPrimeiro.Enabled = Not Edicao '[NAVEGACAO]"
    arrayModuloFuncaoChangeMode(17) = "    btnAnterior.Enabled = Not Edicao '[NAVEGACAO]"
    arrayModuloFuncaoChangeMode(18) = "    btnProximo.Enabled = Not Edicao '[NAVEGACAO]"
    arrayModuloFuncaoChangeMode(19) = "    btnUltimo.Enabled = Not Edicao '[NAVEGACAO]"
    arrayModuloFuncaoChangeMode(20) = "    'os options buttons de operacao"
    arrayModuloFuncaoChangeMode(21) = "    optAlterar.Enabled = Not Edicao"
    arrayModuloFuncaoChangeMode(22) = "    optExcluir.Enabled = Not Edicao"
    arrayModuloFuncaoChangeMode(23) = "    optNovo.Enabled = Not Edicao"
    arrayModuloFuncaoChangeMode(24) = "   "
    arrayModuloFuncaoChangeMode(25) = "    If Not Edicao Then"
    arrayModuloFuncaoChangeMode(26) = "        optAlterar.Value = False"
    arrayModuloFuncaoChangeMode(27) = "        optExcluir.Value = False"
    arrayModuloFuncaoChangeMode(28) = "        optNovo.Value = False"
    arrayModuloFuncaoChangeMode(29) = "        lblStatus.Caption = """""
    arrayModuloFuncaoChangeMode(30) = "    End If"
    arrayModuloFuncaoChangeMode(31) = "    "
    arrayModuloFuncaoChangeMode(32) = "    modoEdicao = Edicao"
    arrayModuloFuncaoChangeMode(33) = "End Sub"
    
    arrayModuloFuncaoCleanControls(1) = "Public Sub CleanControls()"
    arrayModuloFuncaoCleanControls(2) = "        SetValue Me.[NOME_CONTROLE], """""
    arrayModuloFuncaoCleanControls(3) = "End Sub"
    
    arrayModuloFuncaoControlDataType(1) = "Private Function ControlDataType(ctl As MSForms.Control) As String"
    arrayModuloFuncaoControlDataType(2) = "    Select Case ctl.Name"
    arrayModuloFuncaoControlDataType(3) = "    'Case ""txtId"": ControlDataType = ""Integer"""
    arrayModuloFuncaoControlDataType(4) = "    Case ""[NOME_CONTROLE]"": ControlDataType = ""[TIPO_DADO_CONTROLE]"""
    arrayModuloFuncaoControlDataType(5) = "    End Select"
    arrayModuloFuncaoControlDataType(6) = "End Function"
    
    arrayModuloFuncaoSetValues(1) = "Public Sub SetValues(udt[NOME_ENTIDADE] As [NOME_ENTIDADE])"
    arrayModuloFuncaoSetValues(2) = "    With udt[NOME_ENTIDADE]"
    arrayModuloFuncaoSetValues(3) = "        SetValue Me.[NOME_CONTROLE], .[NOME_CAMPO]"
    arrayModuloFuncaoSetValues(4) = "    End With"
    arrayModuloFuncaoSetValues(5) = "    "
    arrayModuloFuncaoSetValues(6) = "    Set cls[NOME_ENTIDADE] = udt[NOME_ENTIDADE]"
    arrayModuloFuncaoSetValues(7) = "End Sub"
    
    arrayModuloFuncaoGetValues(1) = "Public Sub GetValues(ByRef udt[NOME_ENTIDADE] As [NOME_ENTIDADE])"
    arrayModuloFuncaoGetValues(2) = "    With udt[NOME_ENTIDADE]"
    arrayModuloFuncaoGetValues(3) = "        '.Id = GetValue(Me.txtId, TypeName(.Id))"
    arrayModuloFuncaoGetValues(4) = "        .[NOME_CAMPO] = GetValue(Me.[NOME_CONTROLE], TypeName(.[NOME_CAMPO]))"
    arrayModuloFuncaoGetValues(5) = "    End With"
    arrayModuloFuncaoGetValues(6) = "End Sub"
    
    arrayModuloFuncaoQueryClose(1) = "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    arrayModuloFuncaoQueryClose(2) = "    Set cls[NOME_ENTIDADE] = Nothing"
    arrayModuloFuncaoQueryClose(3) = "End Sub"
    
End Sub

Private Sub cboControle_Change()
    If lstColunas.ListIndex > 0 Then
        linha = lstColunas.ListIndex
        lstColunas.List(linha, 1) = cboControle.text
    End If
End Sub

Private Sub cboTabelasFK_Change()
    Dim i As Integer, j As Integer
    
    For i = 1 To UBound(tabelas, 2)
        If tabelas(1, i) = cboTabelasFK.Value Then
            cboFKID.Clear
            cboFKValor.Clear
                        
            For j = 1 To UBound(tabelas(2, i), 2)
                cboFKID.AddItem tabelas(2, i)(colunaCampo, j)
                cboFKValor.AddItem tabelas(2, i)(colunaCampo, j)
            Next j
        End If
    Next i
End Sub

Private Sub cbxFK_Change()
    frameFK.Enabled = cbxFK.Value
    fks(lstColunas.ListIndex + 1, colunaEFK) = IIf(cbxFK.Value, "Sim", "Não")
    
    'força a seleção do ComboBox caso seja selecionada a opção de FK
    If cbxFK.Value Then
        cboControle.Value = "ComboBox"
    End If
End Sub

Private Sub cbxGerar_Click()
    If lstColunas.ListIndex > 0 Then
        linha = lstColunas.ListIndex
        lstColunas.List(linha, 5) = IIf(cbxGerar.Value, "Sim", "Não")
    End If
End Sub

Private Sub cbxNovoArquivo_Click()
    cboArquivosAbertos.Enabled = Not cbxNovoArquivo.Value
    
    If Not cbxNovoArquivo.Value Then
        cboArquivosAbertos.Clear
        Dim plan As Workbook
        For Each plan In Application.Workbooks
            If plan.name <> ThisWorkbook.name Then cboArquivosAbertos.AddItem plan.name
        Next plan
    Else
        cboArquivosAbertos.Clear
    End If
End Sub

Private Sub cbxRequerido_Click()
    If lstColunas.ListIndex > 0 Then
        linha = lstColunas.ListIndex
        lstColunas.List(linha, 2) = IIf(cbxRequerido.Value, "Sim", "Não")
    End If
End Sub

Private Sub cmdConfirmarFK_Click()
    fks(lstColunas.ListIndex + 1, colunaFKTabela) = cboTabelasFK.Value
    fks(lstColunas.ListIndex + 1, colunaFKID) = cboFKID.Value
    fks(lstColunas.ListIndex + 1, colunaFKValor) = cboFKValor.Value
End Sub

Private Sub lstColunas_Click()
    Dim linha As Long
    If lstColunas.ListIndex > 0 Then
        linha = lstColunas.ListIndex
        txtNome.text = lstColunas.List(linha, 0)
        cboControle.text = lstColunas.List(linha, 1)
        cbxRequerido.Value = IIf(lstColunas.List(linha, 2) = "Sim", True, False)
        
        Dim eChave As Boolean
        eChave = (lstColunas.List(linha, 3) = "Sim")
        cboControle.Enabled = Not eChave
        cbxRequerido.Enabled = Not eChave
        cbxGerar.Value = (lstColunas.List(linha, 5) = "Sim")
        
        'campos FK
        cbxFK.Value = IIf(fks(lstColunas.ListIndex + 1, colunaEFK) = "Sim", True, False)
        cboTabelasFK.Value = fks(lstColunas.ListIndex + 1, colunaFKTabela)
        cboFKID.Value = fks(lstColunas.ListIndex + 1, colunaFKID)
        cboFKValor.Value = fks(lstColunas.ListIndex + 1, colunaFKValor)
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
    Dim NomeForm As String
    NomeForm = txtNomeFormulario.text
    If Trim(NomeForm) <> "" Then
        If IsVarArrayEmpty(controles) Then
            MsgBox "E onde estão os campos?"
        Else
            Call CriarForm(Me.txtNomeTabela.text, Trim(RemoveAcentos(NomeForm)))
        End If
    Else
        MsgBox "O nome do formulário é requerido"
    End If
End Sub

Private Sub CriarForm(ByVal NomeEntidade As String, ByVal NomeForm As String)
     
    Dim targetWorkbook As Workbook
    Dim MyUserForm As VBComponent
    Dim nomeEntidadeComAcentos As String
    
    If cbxNovoArquivo.Value Then
        Set targetWorkbook = Application.Workbooks.Add
    Else
        Set targetWorkbook = Application.Workbooks(cboArquivosAbertos.Value)
    End If
    
    nomeEntidadeComAcentos = NomeEntidade
    NomeEntidade = RemoveAcentos(NomeEntidade)
    
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
    Dim nomeRotulo As String
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
    Set modEntidade = targetWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    modEntidade.name = "mod" & NomeEntidade
    
    Call InsertLine(modEntidade, "Sub OpenForm" & NomeForm & "()")
    Call InsertLine(modEntidade, "    'variável do tipo da Classe " & NomeForm)
    Call InsertLine(modEntidade, "    Dim udt" & NomeEntidade & " As " & NomeEntidade)
    Call InsertLine(modEntidade, "    'Cria a isntância")
    Call InsertLine(modEntidade, "    Set udt" & NomeEntidade & " = New " & NomeEntidade)
    Call InsertLine(modEntidade, "    ")
    Call InsertLine(modEntidade, "    udt" & NomeEntidade & ".MoveLast")
    Call InsertLine(modEntidade, "    udt" & NomeEntidade & ".MoveFirst")
    Call InsertLine(modEntidade, "    'Atribui uma instância da classe " & NomeForm & " ao form")
    Call InsertLine(modEntidade, "    ufm" & NomeForm & ".SetValues udt" & NomeEntidade)
    Call InsertLine(modEntidade, "    'Mostra o form")
    Call InsertLine(modEntidade, "    ufm" & NomeForm & ".Show")
    Call InsertLine(modEntidade, "End Sub")
    
    countOfLines = 0
    
    Dim qtdCamposAGerar As Integer
    For i = LBound(controles) To UBound(controles)
        If lstColunas.List(i - 1, colunaGerar - 1) <> "Não" Then qtdCamposAGerar = qtdCamposAGerar + 1
    Next i
    
    'se houver campos a não gerar, aplica a lógica
    If UBound(controles) <> qtdCamposAGerar Then
        'remove os campos que não serão gerados do array
        Dim controlesAGerar()
        
        ReDim controlesAGerar(LBound(controles) To qtdCamposAGerar, LBound(controles, 2) To UBound(controles, 2))
        'copia as colunas
        For i = LBound(controles, 2) To UBound(controles, 2)
            controlesAGerar(1, i) = controles(1, i)
        Next i
        
        'copia o resto
        Dim controlesGerados As Integer
        controlesGerados = 1
        For i = LBound(controles) To UBound(controles)
            If lstColunas.List(i - 1, colunaGerar - 1) <> "Não" Then
                For j = LBound(controles, 2) To UBound(controles, 2)
                    controlesAGerar(controlesGerados, j) = controles(i, j)
                Next j
                controlesGerados = controlesGerados + 1
            End If
        Next i
        
        controles = controlesAGerar
    End If
    
    'gera o modUtilities se ainda não houver
    Dim modUtilitiesExiste As Boolean
    Dim modTypesExiste As Boolean
    
    For N = 1 To targetWorkbook.VBProject.VBComponents.Count
        If targetWorkbook.VBProject.VBComponents(N).name = "modUtilities" Then modUtilitiesExiste = True
        If targetWorkbook.VBProject.VBComponents(N).name = "modTypes" Then modTypesExiste = True
    Next N
    
    If Not modUtilitiesExiste Then
        Dim modUtilities As VBComponent
        Set modUtilities = targetWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
        modUtilities.name = "modUtilities"
        
        Call InsertLine(modUtilities, "Public Function Nz(ByVal Value As Variant) As Variant")
        Call InsertLine(modUtilities, "    If IsNull(Value) Then Value = 0")
        Call InsertLine(modUtilities, "    Nz = Value")
        Call InsertLine(modUtilities, "End Function")
        
        Call InsertLine(modUtilities, "")
        Call InsertLine(modUtilities, "Public Function ObtemCaminhoBancoDeDados() As String")
        Call InsertLine(modUtilities, "    Dim caminhoCompleto As String")
        Call InsertLine(modUtilities, "    Dim ARQUIVO_DADOS As String")
        Call InsertLine(modUtilities, "    Dim PASTA_DADOS As String")
        Call InsertLine(modUtilities, "    ")
        Call InsertLine(modUtilities, "    abrirArquivo = True")
        Call InsertLine(modUtilities, "    ")
        Call InsertLine(modUtilities, "    ARQUIVO_DADOS = ThisWorkbook.Worksheets(""Config"").Range(""ARQUIVO"").Value")
        Call InsertLine(modUtilities, "    PASTA_DADOS = ThisWorkbook.Worksheets(""Config"").Range(""PASTA"").Value")
        Call InsertLine(modUtilities, "    ")
        Call InsertLine(modUtilities, "    If ThisWorkbook.name <> ARQUIVO_DADOS Then")
        Call InsertLine(modUtilities, "        'monta a string do caminho completo")
        Call InsertLine(modUtilities, "        If PASTA_DADOS = vbNullString Or PASTA_DADOS = """" Then")
        Call InsertLine(modUtilities, "            caminhoCompleto = Replace(ThisWorkbook.FullName, ThisWorkbook.name, vbNullString) & ARQUIVO_DADOS")
        Call InsertLine(modUtilities, "        Else")
        Call InsertLine(modUtilities, "            If Right(PASTA_DADOS, 1) = ""\"" Then")
        Call InsertLine(modUtilities, "                caminhoCompleto = PASTA_DADOS & ARQUIVO_DADOS")
        Call InsertLine(modUtilities, "            Else")
        Call InsertLine(modUtilities, "                caminhoCompleto = PASTA_DADOS & ""\"" & ARQUIVO_DADOS")
        Call InsertLine(modUtilities, "            End If")
        Call InsertLine(modUtilities, "        End If")
        Call InsertLine(modUtilities, "    End If")
        Call InsertLine(modUtilities, "    ObtemCaminhoBancoDeDados = caminhoCompleto")
        Call InsertLine(modUtilities, "End Function")
    End If

    countOfLines = 0
    
    Dim modTypes As VBComponent
    If Not modTypesExiste Then
        Set modTypes = targetWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
        modTypes.name = "modTypes"
    Else
        Set modTypes = targetWorkbook.VBProject.VBComponents.Item("modTypes")
    End If
    
    Call InsertLine(modTypes, "Public Type " & NomeEntidade)
    For i = 2 To UBound(controles)
        nomeCampo = RemoveAcentos(CStr(controles(i, colunaCampo)))
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        Call InsertLine(modTypes, "    " & nomeCampo & " As " & tipoDadoControle)
    Next i
    Call InsertLine(modTypes, "End Type")
    Call InsertLine(modTypes, "")
    
    countOfLines = 0
    
    'gera a classe
    Dim classe As VBComponent
    Set classe = targetWorkbook.VBProject.VBComponents.Add(vbext_ct_ClassModule)
    classe.name = NomeEntidade
    
    Call InsertLine(classe, "Private mrstRecordset As Recordset")
    Call InsertLine(classe, "Private mbooLoaded As Boolean")
    Call InsertLine(classe, "Private mdbCurrentDb As Database")
    
    'campos privados
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        nomeControle = ObtemNomeControle(nomeCampo, lstColunas.List(i - 1, colunaControle - 1))
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & RemoveAcentos(nomeCampo)
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
    Call InsertLine(classe, "    If Not Recordset.EOF Then")
    Call InsertLine(classe, "       With Recordset")
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        nomeRotulo = controles(i, colunaRotulo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        If nomeCampo = ChavePrimaria Then
            Call InsertLine(classe, "           " & nomeCampoPrivado & " = Nz(.Fields(""[" & nomeRotulo & "]"").Value)")
        Else
            Call InsertLine(classe, "           Me." & nomeCampo & " = Nz(.Fields(""[" & nomeRotulo & "]"").Value)")
        End If
    Next i
    Call InsertLine(classe, "       End With")
    Call InsertLine(classe, "    Else")
    For i = 2 To UBound(controles)
        nomeCampo = controles(i, colunaCampo)
        nomeRotulo = controles(i, colunaRotulo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        If nomeCampo = ChavePrimaria Then
            Call InsertLine(classe, "           " & nomeCampoPrivado & " = Nz(Empty)")
        Else
            Call InsertLine(classe, "           Me." & nomeCampo & " = Nz(Empty)")
        End If
    Next i
    Call InsertLine(classe, "    End If")
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
        nomeRotulo = controles(i, colunaRotulo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        eRequerido = controles(i, colunaRequerido) = "Sim"
        If nomeCampo = ChavePrimaria Then
            Call InsertLine(classe, "        " & nomeCampoPrivado & " = Nz(.Fields(""[" & nomeRotulo & "]"").Value)")
        Else
            If eRequerido Then
                Call InsertLine(classe, "        .Fields(""[" & nomeRotulo & "]"").Value = NullIfEmptyString(Me." & nomeCampo & ")")
            Else
                Call InsertLine(classe, "        .Fields(""[" & nomeRotulo & "]"").Value = Me." & nomeCampo)
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
    Call InsertLine(classe, "        Set mdbCurrentDb = DBEngine.OpenDatabase(ObtemCaminhoBancoDeDados())")
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
    Call InsertLine(classe, "    If Not Recordset.EOF Then Recordset.MoveFirst")
    Call InsertLine(classe, "    Load")
    Call InsertLine(classe, "End Function")
    Call InsertLine(classe, "")
    Call InsertLine(classe, "'Ocorre quando a classe é instanciada")
    Call InsertLine(classe, "Private Sub Class_Initialize()")
    Call InsertLine(classe, "    Set Recordset = CurrentDb.OpenRecordset(""" & nomeEntidadeComAcentos & """, dbOpenDynaset)")
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
    Call InsertLine(classe, "    If Not Recordset.EOF Then Recordset.MoveLast")
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
     
    'verifica se o formulário existe
    For N = 1 To targetWorkbook.VBProject.VBComponents.Count
        If targetWorkbook.VBProject.VBComponents(N).name = NomeForm Then
            MsgBox "Já existe um formulário com o mesmo nome"
            Exit Sub
        End If
    Next N
    
    Dim alturaForm As Long
    alturaForm = lstColunas.ListCount * 45
    
    If alturaForm < 350 Then alturaForm = 350
    
     
    'Cria o userform
    Set MyUserForm = targetWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    With MyUserForm
        .Properties("Height") = alturaForm
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
        
        If fks(i, colunaEFK) = "Sim" Then
            'ComboBox FK
            Set TextBox = MyUserForm.Designer.Controls.Add("Forms.ComboBox.1")
            With TextBox
                .name = nomeControle
                .Left = margemEsquerda
                .Top = margemTopo + alturaControle + distanciaEntre
                .Height = alturaControle
                .Width = larguraControle
                .ColumnCount = 2
                .ColumnWidths = "0pt;100pt"
            End With
        Else
            'Controle
            Set TextBox = MyUserForm.Designer.Controls.Add("Forms." & tipoControle & ".1")
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
    
    If ckbGerarBotoesDeNavegacao.Value Then
        
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
    End If
    
    'código do form
    With MyUserForm.CodeModule
        countOfLines = .countOfLines
        For i = 1 To UBound(arrayModuloForm)
            Call InsertLine(MyUserForm, ReplaceToken(arrayModuloForm(i)))
        Next i
        
        'função CleanControls
        i = 1
        While i <= UBound(arrayModuloFuncaoCleanControls)
            If InStr(1, arrayModuloFuncaoCleanControls(i), "[NOME_CONTROLE]") > 0 Then
                'guarda a referencia da linha com o conteudo
                linhaNomeControle = i
                For j = 2 To UBound(controles)
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j - 1, colunaControle - 1))
                    linhaAInserir = Replace(arrayModuloFuncaoCleanControls(linhaNomeControle), "[NOME_CONTROLE]", nomeControle)
                    Call InsertLine(MyUserForm, linhaAInserir)
                Next j
            Else
                Call InsertLine(MyUserForm, arrayModuloFuncaoCleanControls(i))
            End If
            i = i + 1
        Wend
        
        'função ChangeMode
        i = 1
        While i <= UBound(arrayModuloFuncaoChangeMode)
            If InStr(1, arrayModuloFuncaoChangeMode(i), "[NOME_CONTROLE]") > 0 Then
                'guarda a referencia da linha com o conteudo
                linhaNomeControle = i
                For j = 2 To UBound(controles)
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j - 1, colunaControle - 1))
                    linhaAInserir = Replace(arrayModuloFuncaoChangeMode(linhaNomeControle), "[NOME_CONTROLE]", nomeControle)
                    Call InsertLine(MyUserForm, ReplaceToken(linhaAInserir))
                Next j
            Else
                If InStr(1, arrayModuloFuncaoChangeMode(i), "[NAVEGACAO]") > 0 Then
                    If ckbGerarBotoesDeNavegacao.Value Then Call InsertLine(MyUserForm, arrayModuloFuncaoChangeMode(i))
                Else
                    Call InsertLine(MyUserForm, ReplaceToken(arrayModuloFuncaoChangeMode(i)))
                End If
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
        
        'função QueryClose
        i = 1
        While i <= UBound(arrayModuloFuncaoQueryClose)
            If InStr(1, arrayModuloFuncaoQueryClose(i), "[NOME_CONTROLE]") > 0 Then
                'guarda a referencia da linha com o conteudo
                linhaNomeControle = i
                For j = 2 To UBound(controles)
                    nomeControle = ObtemNomeControle(controles(j, colunaCampo), lstColunas.List(j - 1, colunaControle - 1))
                    nomeCampo = controles(j, colunaCampo)
                    If nomeCampo <> ChavePrimaria Then
                        linhaAInserir = Replace(arrayModuloFuncaoQueryClose(linhaNomeControle), "[NOME_CONTROLE]", nomeControle)
                        linhaAInserir = Replace(linhaAInserir, "[NOME_CAMPO]", nomeCampo)
                        Call InsertLine(MyUserForm, linhaAInserir)
                    End If
                Next j
            Else
                Call InsertLine(MyUserForm, ReplaceToken(arrayModuloFuncaoQueryClose(i)))
            End If
            i = i + 1
        Wend
        
        'LoadDependentCombobox
        Call InsertLine(MyUserForm, "Private Sub LoadDependentCombos()")
        Call InsertLine(MyUserForm, "   Dim rstFiltro As Recordset")
        Call InsertLine(MyUserForm, "   Dim arrayItems() As String")
        Call InsertLine(MyUserForm, "   Dim filtros As String")
        Call InsertLine(MyUserForm, "   Dim linha As Long")
    
        For i = 2 To UBound(fks, 1)
            If fks(i, colunaEFK) = "Sim" Then
                Call InsertLine(MyUserForm, "   '" & fks(i, colunaFKTabela))
                Call InsertLine(MyUserForm, "   Dim cls" & fks(i, colunaFKTabela) & " As " & fks(i, colunaFKTabela))
                Call InsertLine(MyUserForm, "   Set cls" & fks(i, colunaFKTabela) & " = New " & fks(i, colunaFKTabela))
                Call InsertLine(MyUserForm, "   ")
                Call InsertLine(MyUserForm, "   Me.cbo" & fks(i, colunaFKCampo) & ".Clear")
                Call InsertLine(MyUserForm, "   'filtros = filtros & ""[ATIVO] = True""")
                Call InsertLine(MyUserForm, "   cls" & fks(i, colunaFKTabela) & ".Filter = filtros")
                Call InsertLine(MyUserForm, "   Set rstFiltro = cls" & fks(i, colunaFKTabela) & ".RecordSetWithFilter")
                Call InsertLine(MyUserForm, "   ")
                Call InsertLine(MyUserForm, "   ReDim arrayItems(1 To rstFiltro.RecordCount, 1 To 2)")
                Call InsertLine(MyUserForm, "   ")
                Call InsertLine(MyUserForm, "   linha = 1")
                Call InsertLine(MyUserForm, "   'linhas")
                Call InsertLine(MyUserForm, "   Do While Not rstFiltro.EOF")
                Call InsertLine(MyUserForm, "       arrayItems(linha, 1) = rstFiltro(""" & fks(i, colunaFKID) & """)")
                Call InsertLine(MyUserForm, "       arrayItems(linha, 2) = rstFiltro(""" & fks(i, colunaFKValor) & """)")
                Call InsertLine(MyUserForm, "       linha = linha + 1")
                Call InsertLine(MyUserForm, "       rstFiltro.MoveNext")
                Call InsertLine(MyUserForm, "   Loop")
                Call InsertLine(MyUserForm, "   ")
                Call InsertLine(MyUserForm, "   Me.cbo" & fks(i, colunaFKCampo) & ".List = arrayItems")
                Call InsertLine(MyUserForm, "   ")
                Call InsertLine(MyUserForm, "   rstFiltro.Close")
                Call InsertLine(MyUserForm, "   Set rstFiltro = Nothing")
                Call InsertLine(MyUserForm, "   Set cls" & fks(i, colunaFKTabela) & " = Nothing")
                Call InsertLine(MyUserForm, "   Erase arrayItems")
                Call InsertLine(MyUserForm, "   filtros = """"")
            End If
        Next i
        Call InsertLine(MyUserForm, "End Sub")
    End With
    
    'Formulário de Pesquisa
    Dim NomeFormPesquisa As String
    Dim UserFormPesquisa As VBComponent
    NomeFormPesquisa = NomeForm & "Pesquisa"
     
    'verifica se o formulário exite
    For N = 1 To targetWorkbook.VBProject.VBComponents.Count
        If targetWorkbook.VBProject.VBComponents(N).name = NomeFormPesquisa Then
            MsgBox "Já existe um formulário com o mesmo nome"
            Exit Sub
        End If
    Next N
     
    alturaForm = lstColunas.ListCount * 45
    'Cria o userform
    Set UserFormPesquisa = targetWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
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
        nomeRotulo = controles(i, colunaRotulo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        eRequerido = controles(i, colunaRequerido) = "Sim"
        nomeControle = ObtemNomeControle(nomeCampo, lstColunas.List(i - 1, colunaControle - 1))
        tipoControle = lstColunas.List(i - 1, colunaControle - 1)
        
        If tipoControle = "CheckBox" Then
            'Checkbox de filtro
            Set CheckBox = UserFormPesquisa.Designer.Controls.Add("Forms.CheckBox.1")
            With CheckBox
                .Caption = nomeRotulo
                .name = nomeControle
                .Left = margemEsquerda
                .Top = margemTopo
                .Height = alturaControle
                .Width = larguraControle / 2
            End With
            'Checkbox de controle de filtro
            Set CheckBoxFiltro = UserFormPesquisa.Designer.Controls.Add("Forms.CheckBox.1")
            With CheckBoxFiltro
                .Caption = "Filtra " & nomeRotulo
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
                .Caption = nomeRotulo
                .name = "lbl" & controles(i, colunaCampo)
                .Left = margemEsquerda
                .Top = margemTopo
                .Height = alturaControle
                .Width = larguraControle
            End With
            If fks(i, colunaEFK) = "Sim" Then
                'ComboBox FK
                Set TextBox = UserFormPesquisa.Designer.Controls.Add("Forms.ComboBox.1")
                With TextBox
                    .name = nomeControle
                    .Left = margemEsquerda
                    .Top = margemTopo + alturaControle + distanciaEntre
                    .Height = alturaControle
                    .Width = larguraControle
                    .ColumnCount = 2
                    .ColumnWidths = "0pt;100pt"
                End With
            Else
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
    For i = 2 To UBound(fks, 2)
        If fks(i, colunaEFK) = "Sim" Then
            Call InsertLine(UserFormPesquisa, "Private array" & fks(i, colunaFKTabela) & "() As String")
        End If
    Next i
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub btnCancelar_Click()")
    Call InsertLine(UserFormPesquisa, "    Unload Me")
    Call InsertLine(UserFormPesquisa, "End Sub")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub btnOK_Click()")
    Call InsertLine(UserFormPesquisa, "    Call FillListBox")
    Call InsertLine(UserFormPesquisa, "End Sub")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub lst" & NomeEntidade & "_DblClick(ByVal Cancel As MSForms.ReturnBoolean)")
    Call InsertLine(UserFormPesquisa, "    If lst" & NomeEntidade & ".ListIndex > 0 Then")
    Call InsertLine(UserFormPesquisa, "        Dim " & ChavePrimaria)
    Call InsertLine(UserFormPesquisa, "        " & ChavePrimaria & " = lst" & NomeEntidade & ".List(lst" & NomeEntidade & ".ListIndex, 0)")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "        cls" & NomeEntidade & ".MoveFirst")
    Call InsertLine(UserFormPesquisa, "        Do")
    Call InsertLine(UserFormPesquisa, "            If cls" & NomeEntidade & "." & ChavePrimaria & " = " & ChavePrimaria & " Then")
    Call InsertLine(UserFormPesquisa, "                " & NomeForm & ".SetValues cls" & NomeEntidade & "")
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
    Call InsertLine(UserFormPesquisa, "    Call LoadDependentCombos")
    Call InsertLine(UserFormPesquisa, "    Call FillListBox")
    Call InsertLine(UserFormPesquisa, "End Sub")
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub FillListBox()")
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
        nomeRotulo = controles(i, colunaRotulo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        nomeCampoPrivado = "m" & ObtemAcronimoTipo(tipoDadoControle) & nomeCampo
        eRequerido = controles(i, colunaRequerido) = "Sim"
        nomeControle = ObtemNomeControle(nomeCampo, lstColunas.List(i - 1, colunaControle - 1))
        tipoControle = lstColunas.List(i - 1, colunaControle - 1)
        
        If nomeCampo = ChavePrimaria Then
            Call InsertLine(UserFormPesquisa, "    If Trim(" & nomeControle & ".Text) <> """" Then")
            Call InsertLine(UserFormPesquisa, "        filtros = ""[" & nomeRotulo & "] = "" & Trim(" & nomeControle & ".Text)")
            Call InsertLine(UserFormPesquisa, "    End If")
            Call InsertLine(UserFormPesquisa, "")
        Else
            If tipoDadoControle = "String" Then
                Call InsertLine(UserFormPesquisa, "    If Trim(" & nomeControle & ".Text) <> """" Then")
                Call InsertLine(UserFormPesquisa, "        If filtros <> """" Then filtros = filtros & "" AND """)
                If tipoControle = "ComboBox" Then
                    Call InsertLine(UserFormPesquisa, "        filtros = filtros & ""[" & nomeRotulo & "] LIKE '*"" & Trim(" & nomeControle & ".Value) & ""*'""")
                Else
                    Call InsertLine(UserFormPesquisa, "        filtros = filtros & ""[" & nomeRotulo & "] LIKE '*"" & Trim(" & nomeControle & ".Text) & ""*'""")
                End If
                Call InsertLine(UserFormPesquisa, "    End If")
                Call InsertLine(UserFormPesquisa, "")
            ElseIf tipoDadoControle = "Date" Then
                Call InsertLine(UserFormPesquisa, "    If Trim(" & nomeControle & ".Text) <> """" Then")
                Call InsertLine(UserFormPesquisa, "        If filtros <> """" Then filtros = filtros & "" AND """)
                Call InsertLine(UserFormPesquisa, "        filtros = filtros & ""[" & nomeRotulo & "] = #"" & Trim(CDate(" & nomeControle & ".Text)) & ""#""")
                Call InsertLine(UserFormPesquisa, "    End If")
                Call InsertLine(UserFormPesquisa, "")
            ElseIf tipoDadoControle = "Boolean" Then
                If tipoControle = "CheckBox" Then Call InsertLine(UserFormPesquisa, "    If " & nomeControle & "Filtrar.Value Then")
                Call InsertLine(UserFormPesquisa, "        If " & nomeControle & ".Value <> """" Then")
                Call InsertLine(UserFormPesquisa, "            If filtros <> """" Then filtros = filtros & "" AND """)
                Call InsertLine(UserFormPesquisa, "            filtros = filtros & ""[" & nomeRotulo & "] = "" & IIf(" & nomeControle & ".Value, ""True"", ""False"")")
                Call InsertLine(UserFormPesquisa, "        End If")
                If tipoControle = "CheckBox" Then Call InsertLine(UserFormPesquisa, "    End If")
            Else
                Call InsertLine(UserFormPesquisa, "    If Trim(" & nomeControle & ".Text) <> """" Then")
                Call InsertLine(UserFormPesquisa, "        If filtros <> """" Then filtros = filtros & "" AND """)
                Call InsertLine(UserFormPesquisa, "        filtros = filtros & ""[" & nomeRotulo & "] = "" & Trim(" & nomeControle & ".Text)")
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
        nomeRotulo = controles(i, colunaRotulo)
        tipoDadoControle = ObtemTipoDadoCampo(controles(i, colunaControle))
        If tipoDadoControle = "Boolean" Then
            Call InsertLine(UserFormPesquisa, "        arrayItems(linha, " & i - 1 & ") = IIf(rstFiltro(""[" & nomeCampo & "]"") ,""Sim"",""Não"")")
        Else
            If fks(i, colunaEFK) = "Sim" Then
                Call InsertLine(UserFormPesquisa, "        arrayItems(linha, " & i - 1 & ") = LookUpArray(rstFiltro(""[" & nomeCampo & "]""), array" & fks(i, colunaFKTabela) & ")")
            Else
                Call InsertLine(UserFormPesquisa, "        arrayItems(linha, " & i - 1 & ") = rstFiltro(""[" & nomeCampo & "]"")")
            End If
        End If
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
    Call InsertLine(UserFormPesquisa, "")
    Call InsertLine(UserFormPesquisa, "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)")
    Call InsertLine(UserFormPesquisa, "    Set cls" & NomeEntidade & " = Nothing")
    For i = 2 To UBound(fks, 2)
        If fks(i, colunaEFK) = "Sim" Then
            Call InsertLine(UserFormPesquisa, "    Erase array" & fks(i, colunaFKTabela))
        End If
    Next i
    Call InsertLine(UserFormPesquisa, "End Sub")
    
    Call InsertLine(UserFormPesquisa, "Public Function LookUpArray(ByVal LookUpValue As String, ByRef arrayItems() As String) As String")
    Call InsertLine(UserFormPesquisa, "    Dim FoundValue As String")
    Call InsertLine(UserFormPesquisa, "    ")
    Call InsertLine(UserFormPesquisa, "    For i = 1 To UBound(arrayItems, 1)")
    Call InsertLine(UserFormPesquisa, "        If arrayItems(i, 1) = LookUpValue Then")
    Call InsertLine(UserFormPesquisa, "            FoundValue = arrayItems(i, 2)")
    Call InsertLine(UserFormPesquisa, "            Exit For")
    Call InsertLine(UserFormPesquisa, "        End If")
    Call InsertLine(UserFormPesquisa, "    Next i")
    Call InsertLine(UserFormPesquisa, "    ")
    Call InsertLine(UserFormPesquisa, "    LookUpArray = FoundValue")
    Call InsertLine(UserFormPesquisa, "End Function")
    
    'LoadDependentCombobox
    Call InsertLine(UserFormPesquisa, "Private Sub LoadDependentCombos()")
    Call InsertLine(UserFormPesquisa, "   Dim rstFiltro As Recordset")
    Call InsertLine(UserFormPesquisa, "   Dim arrayItems() As String")
    Call InsertLine(UserFormPesquisa, "   Dim filtros As String")
    Call InsertLine(UserFormPesquisa, "   Dim linha As Long")
    For i = 2 To UBound(fks, 1)
        If fks(i, colunaEFK) = "Sim" Then
            Call InsertLine(UserFormPesquisa, "   '" & fks(i, colunaFKTabela))
            Call InsertLine(UserFormPesquisa, "   Dim cls" & fks(i, colunaFKTabela) & " As " & fks(i, colunaFKTabela))
            Call InsertLine(UserFormPesquisa, "   Set cls" & fks(i, colunaFKTabela) & " = New " & fks(i, colunaFKTabela))
            Call InsertLine(UserFormPesquisa, "   ")
            Call InsertLine(UserFormPesquisa, "   Me.cbo" & fks(i, colunaFKCampo) & ".Clear")
            Call InsertLine(UserFormPesquisa, "   'filtros = filtros & ""[ATIVO] = True""")
            Call InsertLine(UserFormPesquisa, "   cls" & fks(i, colunaFKTabela) & ".Filter = filtros")
            Call InsertLine(UserFormPesquisa, "   Set rstFiltro = cls" & fks(i, colunaFKTabela) & ".RecordSetWithFilter")
            Call InsertLine(UserFormPesquisa, "   ")
            Call InsertLine(UserFormPesquisa, "   ReDim arrayItems(1 To rstFiltro.RecordCount, 1 To 2)")
            Call InsertLine(UserFormPesquisa, "   ")
            Call InsertLine(UserFormPesquisa, "   linha = 1")
            Call InsertLine(UserFormPesquisa, "   'linhas")
            Call InsertLine(UserFormPesquisa, "   Do While Not rstFiltro.EOF")
            Call InsertLine(UserFormPesquisa, "       arrayItems(linha, 1) = rstFiltro(""" & fks(i, colunaFKID) & """)")
            Call InsertLine(UserFormPesquisa, "       arrayItems(linha, 2) = rstFiltro(""" & fks(i, colunaFKValor) & """)")
            Call InsertLine(UserFormPesquisa, "       linha = linha + 1")
            Call InsertLine(UserFormPesquisa, "       rstFiltro.MoveNext")
            Call InsertLine(UserFormPesquisa, "   Loop")
            Call InsertLine(UserFormPesquisa, "   ")
            Call InsertLine(UserFormPesquisa, "   Me.cbo" & fks(i, colunaFKCampo) & ".List = arrayItems")
            Call InsertLine(UserFormPesquisa, "   ")
            Call InsertLine(UserFormPesquisa, "   rstFiltro.Close")
            Call InsertLine(UserFormPesquisa, "   Set rstFiltro = Nothing")
            Call InsertLine(UserFormPesquisa, "   Set cls" & fks(i, colunaFKTabela) & " = Nothing")
            Call InsertLine(UserFormPesquisa, "   array" & fks(i, colunaFKTabela) & " = arrayItems")
            Call InsertLine(UserFormPesquisa, "   Erase arrayItems")
            Call InsertLine(UserFormPesquisa, "   filtros = """"")
        End If
    Next i
    Call InsertLine(UserFormPesquisa, "End Sub")
    
    'Adiciona as referencias no novo arquivo
    Dim ref As Reference
    For Each ref In ThisWorkbook.VBProject.References
        Call targetWorkbook.VBProject.References.AddFromGuid(ref.GUID, ref.Major, ref.Minor)
    Next ref
    
    'Define a planilha de configuração e valores
    Dim ws As Worksheet, rngPasta As Range, rngArquivo As Range
    Set ws = targetWorkbook.Worksheets(1)
    ws.name = "Config"
    Call targetWorkbook.Names.Add("PASTA", "=Config!R1C2")
    Call targetWorkbook.Names.Add("ARQUIVO", "=Config!R2C2")
    ws.Cells(1, 1).Value = "Pasta:"
    ws.Cells(2, 1).Value = "Arquivo:"
    Dim arrayArquivo() As String
    arrayArquivo = Split(ufmSelecionaBanco.txtCaminhoBanco.text, "\")
    ws.Cells(2, 2).Value = arrayArquivo(UBound(arrayArquivo))
    ws.Cells(1, 2).Value = Replace(ufmSelecionaBanco.txtCaminhoBanco.text, ws.Cells(2, 2).Value, "", Compare:=vbTextCompare)
    
    Debug.Print "CountOfLines :" & countOfLines
    
    MsgBox NomeForm & " gerado com sucesso"
    Unload Me
    
    If MsgBox("Gerar mais formulários do mesmo banco de dados?", vbYesNo) = vbNo Then Unload ufmSelecionaBanco
End Sub

Private Sub InsertLine(ByRef componente As VBComponent, ByVal linha As String)
    countOfLines = countOfLines + 1
    Call componente.CodeModule.InsertLines(countOfLines, linha)
    'Debug.Print Linha
End Sub

Private Function ReplaceToken(ByVal text As String)
    Dim i As Integer
    '[NOME_ENTIDADE]
    text = Replace(text, "[NOME_ENTIDADE]", Trim(RemoveAcentos(txtNomeTabela.text)))
    '[NOME_FORM]
    text = Replace(text, "[NOME_FORM]", Trim(RemoveAcentos(txtNomeFormulario.text)))
    '[CHAVE_PRIMARIA]
    text = Replace(text, "[CHAVE_PRIMARIA]", ChavePrimaria)
    '[CONTROLES_REQUERIDOS]
    Dim controlesRequeridos() As String, controlesRequeridosCount As Long, controlesRequeridosIndex As Long
    i = 1
    Do
        If controles(i, colunaRequerido) = "Sim" Then controlesRequeridosCount = controlesRequeridosCount + 1
        i = i + 1
    Loop While i <= UBound(controles)
        
    If controlesRequeridosCount > 0 Then
        ReDim controlesRequeridos(1 To controlesRequeridosCount)
        controlesRequeridosIndex = 1
        For i = 2 To UBound(controles)
            If controles(i, colunaRequerido) = "Sim" Then
                controlesRequeridos(controlesRequeridosIndex) = """" & ObtemNomeControle(controles(i, colunaCampo), lstColunas.List(i - 1, colunaControle - 1)) & """"
                controlesRequeridosIndex = controlesRequeridosIndex + 1
            End If
        Next i
        text = Replace(text, "[CONTROLES_REQUERIDOS]", Join(controlesRequeridos, ","))
    Else
        text = Replace(text, "[CONTROLES_REQUERIDOS]", """")
    End If
    
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Erase tabelas
    Erase controles
    Erase fks
    Erase arrayModuloForm
    Erase arrayModuloFuncaoCleanControls
    Erase arrayModuloFuncaoControlDataType
    Erase arrayModuloFuncaoSetValues
    Erase arrayModuloFuncaoGetValues
    Erase arrayModuloFuncaoQueryClose
    Erase arrayModuloFuncaoChangeMode
End Sub

