Attribute VB_Name = "modAccess"
Option Explicit

Private mdbCurrentDb As Database
Private path As String

Public Function CurrentDb(ByVal caminhoArquivo As String) As Database
    If path <> caminhoArquivo Then Set mdbCurrentDb = Nothing
    
    If mdbCurrentDb Is Nothing Then
        If caminhoArquivo = "" Then
            caminhoArquivo = OpenFileDialog
        Else
            Set mdbCurrentDb = DBEngine.OpenDatabase(caminhoArquivo)
        End If
    End If

    path = caminhoArquivo
    Set CurrentDb = mdbCurrentDb
End Function

Public Sub CloseConnection()
    If Not mdbCurrentDb Is Nothing Then
        Call mdbCurrentDb.Close
        Set mdbCurrentDb = Nothing
    End If
End Sub

Public Function OpenFileDialog() As String
    Dim Filter As String, Title As String
    Dim FilterIndex As Integer
    Dim fileName As Variant
    ' Define o filtro de procura dos arquivos
    Filter = "Microsoft Access (*.accdb),*.mdb,"
    ' O filtro padrão é *.*
    FilterIndex = 3
    ' Define o Título (Caption) da Tela
    Title = "Selecione um arquivo"
    ' Define o disco de procura
    ChDrive ("C")
    ChDir ("C:\")
    With Application
        ' Abre a caixa de diálogo para seleção do arquivo com os parâmetros
        fileName = .GetOpenFilename(Filter, FilterIndex, Title)
        ' Reseta o Path
        ChDrive (Left(.DefaultFilePath, 1))
        ChDir (.DefaultFilePath)
    End With
    ' Abandona ao Cancelar
    If fileName = False Then
        MsgBox "Nenhum arquivo foi selecionado."
        Exit Function
    End If
    ' Retorna o caminho do arquivo
    OpenFileDialog = fileName
End Function

 
Public Sub ListTablesAndFields(ByVal caminhoArquivo As String, ByRef tabelas())
     'Macro Purpose:  Write all table and field names to and Excel file
     
    Dim lTbl As Long
    Dim lFld As Long
    Dim dBase As Database
    Dim campos()
     
     'Set current database to a variable adn create a new Excel instance
    Set dBase = CurrentDb(caminhoArquivo)
     
     'Set on error in case there is no tables
    On Error Resume Next
     
    'ReDim tabelas(1 To 2, 1 To dBase.TableDefs.Count)
    Dim totalTabelas As Integer
     'Loop through all tables
    For lTbl = 0 To dBase.TableDefs.Count
         'If the table name is a temporary or system table then ignore it
        If Left(dBase.TableDefs(lTbl).name, 1) = "~" Or UCase(Left(dBase.TableDefs(lTbl).name, 4)) = "MSYS" Then
             '~ indicates a temporary table
             'MSYS indicates a system level table
        Else
            Dim primaryKey As String
            primaryKey = dBase.TableDefs(lTbl).Indexes("PrimaryKey").Fields(0).name
            totalTabelas = totalTabelas + 1
            ReDim Preserve tabelas(1 To 2, 1 To totalTabelas)
            tabelas(1, totalTabelas) = dBase.TableDefs(lTbl).name
             'Otherwise, loop through each table, writing the table and field names
             'to the Excel file
            Dim fieldCount As Long
            Erase campos
            ReDim campos(1 To 5, 1 To dBase.TableDefs(lTbl).Fields.Count)
            For lFld = 1 To dBase.TableDefs(lTbl).Fields.Count
                campos(1, lFld) = RemoveAcentos(dBase.TableDefs(lTbl).Fields(lFld - 1).name)
                campos(2, lFld) = FieldTypeName(dBase.TableDefs(lTbl).Fields(lFld - 1))
                campos(3, lFld) = IIf(dBase.TableDefs(lTbl).Fields(lFld - 1).Required, "Sim", "Não")
                campos(4, lFld) = IIf(dBase.TableDefs(lTbl).Fields(lFld - 1).name = primaryKey, "Sim", "Não")
                campos(5, lFld) = dBase.TableDefs(lTbl).Fields(lFld - 1).name
            Next lFld
            tabelas(2, totalTabelas) = campos
        End If
    Next lTbl
     'Resume error breaks
    On Error GoTo 0
     
     'Release database object from memory
    Set dBase = Nothing
End Sub

Function FieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function
