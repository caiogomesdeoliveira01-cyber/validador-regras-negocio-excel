Attribute VB_Name = "limpar_tabela"
Option Explicit

Sub limpar_tabela()

    Dim tbl As ListObject

    On Error GoTo Erro

    Application.ScreenUpdating = False
    Set tbl = ThisWorkbook.Worksheets("Tabela").ListObjects("tbl_tabela")

    LimparDadosTabela tbl

    Application.ScreenUpdating = True
    MsgBox "Tabela limpa com sucesso!", vbInformation
    Exit Sub

Erro:
    Application.ScreenUpdating = True
    MsgBox "Erro ao limpar a tabela:" & vbCrLf & Err.Description, vbCritical

End Sub


Sub excluir_tabela()

    Dim tbl As ListObject
    Dim i As Long

    On Error GoTo Erro

    Application.ScreenUpdating = False
    Set tbl = ThisWorkbook.Worksheets("Tabela").ListObjects("tbl_tabela")

    'Remove filtros ativos para evitar erros de exclusão
    RemoverFiltros tbl

    'Exclui todas as colunas, mantendo apenas a primeira (de trás para frente)
    For i = tbl.ListColumns.Count To 2 Step -1
        tbl.ListColumns(i).Delete
    Next i

    LimparDadosTabela tbl

    tbl.ListColumns(1).Name = "Cole a tabela aqui"
    tbl.Range.Columns.AutoFit

    Application.ScreenUpdating = True
    MsgBox "Tabela restaurada com sucesso!", vbInformation
    Exit Sub

Erro:
    Application.ScreenUpdating = True
    MsgBox "Erro ao restaurar a tabela:" & vbCrLf & Err.Description, vbCritical

End Sub


'--------------------------------------------------------------
' Rotinas auxiliares
'--------------------------------------------------------------
Private Sub RemoverFiltros(ByVal tbl As ListObject)

    If tbl.ShowAutoFilter Then
        If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    End If

End Sub


Private Sub LimparDadosTabela(ByVal tbl As ListObject)

    RemoverFiltros tbl

    'Garante que exista ao menos 1 linha de dados
    If tbl.DataBodyRange Is Nothing Then tbl.ListRows.Add

    'Apaga da 2ª linha em diante (a tabela nunca fica vazia)
    If tbl.ListRows.Count > 1 Then
        tbl.DataBodyRange.Offset(1, 0).Resize(tbl.ListRows.Count - 1).Delete
    End If

    'Limpa o conteúdo da única linha que sobrou
    tbl.DataBodyRange.ClearContents

End Sub
