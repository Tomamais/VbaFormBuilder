VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmSelecionaBanco 
   Caption         =   "Selecionar o banco de dados"
   ClientHeight    =   4380
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6456
   OleObjectBlob   =   "ufmSelecionaBanco.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufmSelecionaBanco"
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
Private Sub btnConfigurarCampos_Click()
    If lstCampos.ListCount > 0 Then
        ufmConstrutor.lstColunas.ColumnCount = lstCampos.ColumnCount
        ufmConstrutor.lstColunas.List = lstCampos.List
        'substitui a coluna tipo por controle
        Dim i As Integer
        ufmConstrutor.lstColunas.List(1, 1) = "Controle"
        For i = 1 To ufmConstrutor.lstColunas.ListCount - 1
            ufmConstrutor.lstColunas.List(i, 1) = "TextBox"
        Next i
        ufmConstrutor.txtNomeFormulario.text = cboTabelas.text
        Call ufmConstrutor.DefineControles(controles)
        ufmConstrutor.Show
    Else
        MsgBox "Não há campos para gerar o form"
    End If
End Sub

Private Sub btnSelecionarArquivo_Click()
    Dim recarrega As Boolean
    If caminhoArquivo = "" Then
        caminhoArquivo = OpenFileDialog
        recarrega = True
    Else
        Dim tempCaminho As String
        tempCaminho = OpenFileDialog
        If tempCaminho <> caminhoArquivo Then
            recarrega = True
        End If
    End If
    
    If recarrega And caminhoArquivo <> "" Then
        Call ListTablesAndFields(caminhoArquivo, tabelas)
        Dim i As Integer
        For i = 1 To UBound(tabelas, 2)
            cboTabelas.AddItem tabelas(1, i)
        Next i
    End If
    
    txtCaminhoBanco.text = caminhoArquivo
End Sub

Private Sub cboTabelas_Change()
    Dim i As Integer, j As Integer
    lstCampos.ColumnCount = 4
    lstCampos.Clear
    
    For i = 1 To UBound(tabelas, 2)
        If tabelas(1, i) = cboTabelas.text Then
            ReDim controles(1 To UBound(tabelas(2, i), 2) + 1, 1 To 4)
                        
            controles(1, colunaCampo) = "Campo"
            controles(1, colunaTipo) = "Tipo"
            controles(1, colunaRequerido) = "Requerido"
            controles(1, colunaEchave) = "É chave?"
            
            For j = 1 To UBound(tabelas(2, i), 2)
                controles(j + 1, colunaCampo) = tabelas(2, i)(colunaCampo, j)
                controles(j + 1, colunaTipo) = tabelas(2, i)(colunaTipo, j)
                controles(j + 1, colunaRequerido) = tabelas(2, i)(colunaRequerido, j)
                controles(j + 1, colunaEchave) = tabelas(2, i)(colunaEchave, j)
            Next j
        End If
    Next i
    
    lstCampos.List = controles
    
    
'    For i = 1 To UBound(tabelas, 2)
'        If tabelas(1, i) = cboTabelas.Text Then
'            lstCampos.AddItem ""
'            lstCampos.List(0, 0) = "Campo"
'            lstCampos.List(0, 1) = "Tipo"
'            lstCampos.List(0, 2) = "Requerido"
'            lstCampos.List(0, 3) = "É chave"
'            For j = 1 To UBound(tabelas(2, i), 2)
'                lstCampos.AddItem ""
'                lstCampos.List(j, 0) = tabelas(2, i)(1, j)
'                lstCampos.List(j, 1) = tabelas(2, i)(2, j)
'                lstCampos.List(j, 2) = tabelas(2, i)(3, j)
'                lstCampos.List(j, 3) = tabelas(2, i)(4, j)
'            Next j
'        End If
'    Next
End Sub

Private Sub UserForm_Click()

End Sub
