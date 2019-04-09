VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmSelecionaBancoMultiplo 
   Caption         =   "Selecionar o banco de dados"
   ClientHeight    =   6936
   ClientLeft      =   96
   ClientTop       =   372
   ClientWidth     =   8940.001
   OleObjectBlob   =   "ufmSelecionaBancoMultiplo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufmSelecionaBancoMultiplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private caminhoArquivo As String
Private tabelas()
Private controles()
Private Const colunaCampo As Integer = 1
Private Const colunaTipo As Integer = 2
Private Const colunaRequerido As Integer = 3
Private Const colunaEchave As Integer = 4
Private Const colunaRotulo As Integer = 5
Private Const colunaGerar As Integer = 6

Private Sub btnConfigurarCampos_Click()
    Dim i As Integer, temChave As Boolean
    'só segue em frente se tiver pelo menos um campo de chave primária
    For i = 2 To UBound(controles)
        If controles(i, colunaEchave) = "Sim" Then
            temChave = True
            Exit For
        End If
    Next i
    
    If Not temChave Then
        MsgBox "A tabela precisa ter pelo menos uma chave primária"
        Exit Sub
    End If

    If lstCampos.ListCount > 0 Then
        ufmConstrutor.lstColunas.ColumnCount = lstCampos.ColumnCount
        ufmConstrutor.lstColunas.List = lstCampos.List
        'substitui a coluna tipo por controle
        ufmConstrutor.lstColunas.List(1, 1) = "Controle"
        For i = 1 To ufmConstrutor.lstColunas.ListCount - 1
            ufmConstrutor.lstColunas.List(i, 1) = "TextBox"
        Next i
        ufmConstrutor.txtNomeFormulario.text = cboTabelas.text
        ufmConstrutor.txtNomeTabela.text = cboTabelas.text
        Call ufmConstrutor.DefineControles(controles)
        Call ufmConstrutor.DefineTabelas(tabelas)
        ufmConstrutor.Show
    Else
        MsgBox "Não há campos para gerar o form"
    End If
End Sub

Private Sub btnSelecionarArquivo_Click()
    Dim i As Integer
    Dim recarrega As Boolean
    If caminhoArquivo = "" Then
        caminhoArquivo = OpenFileDialog
        recarrega = True
    Else
        Dim tempCaminho As String
        tempCaminho = OpenFileDialog
        If tempCaminho <> caminhoArquivo Then
            caminhoArquivo = tempCaminho
            recarrega = True
        End If
    End If
    
    If recarrega And caminhoArquivo <> "" Then
        Erase tabelas
        Call ListTablesAndFields(caminhoArquivo, tabelas)
        cboTabelas.Clear
        For i = 1 To UBound(tabelas, 2)
            cboTabelas.AddItem tabelas(1, i)
        Next i
        
        Dim tabelasToListBox As Variant
        tabelasToListBox = Array2DTranspose(tabelas)
        lstTabelas.List = tabelasToListBox
    End If
    
    txtCaminhoBanco.text = caminhoArquivo
End Sub

Private Sub cboTabelas_Change()
    Dim i As Integer, j As Integer
    lstCampos.ColumnCount = 6
    lstCampos.Clear
    
    For i = 1 To UBound(tabelas, 2)
        If tabelas(1, i) = cboTabelas.text Then
            Erase controles
            ReDim controles(1 To UBound(tabelas(2, i), 2) + 1, 1 To 6)
                        
            controles(1, colunaCampo) = "Campo"
            controles(1, colunaTipo) = "Tipo"
            controles(1, colunaRequerido) = "Requerido"
            controles(1, colunaEchave) = "É chave?"
            controles(1, colunaRotulo) = "Rótulo"
            controles(1, colunaGerar) = "Gerar?"
            
            For j = 1 To UBound(tabelas(2, i), 2)
                controles(j + 1, colunaCampo) = tabelas(2, i)(colunaCampo, j)
                controles(j + 1, colunaTipo) = tabelas(2, i)(colunaTipo, j)
                controles(j + 1, colunaRequerido) = tabelas(2, i)(colunaRequerido, j)
                controles(j + 1, colunaEchave) = tabelas(2, i)(colunaEchave, j)
                controles(j + 1, colunaRotulo) = tabelas(2, i)(colunaRotulo, j)
                controles(j + 1, colunaGerar) = "Sim"
            Next j
        End If
    Next i
    
    lstCampos.List = controles
End Sub

Private Sub cmdLimparSelecao_Click()
    If lstTabelas.ListCount > 0 Then
        For i = 0 To lstTabelas.ListCount - 1
            lstTabelas.Selected(i) = False
        Next i
    End If
End Sub

Private Sub cmdSelecionarTodos_Click()
    If lstTabelas.ListCount > 0 Then
        For i = 0 To lstTabelas.ListCount - 1
            lstTabelas.Selected(i) = True
        Next i
    End If
End Sub

Private Sub lstTabelas_Change()
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call CloseConnection
End Sub
