Attribute VB_Name = "Module1"
Sub file_transfer()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim srcAddress As String
    Dim srcFileName As String
    Dim destAddress As String
    Dim destFileName As String
    Dim processType As String
    Dim status As String
    Dim executeTime As String
    Dim errorMessage As String
    
    ' シートを取得（シート名は必要に応じて変更してください。）
    Set ws = ThisWorkbook.Sheets("メイン")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 処理内容ごとに処理を行う
    For i = 2 To lastRow
        ' 必須項目の取得
        processType = ws.Cells(i, 1).Value
        srcAddress = ws.Cells(i, 2).Value
        srcFileName = ws.Cells(i, 3).Value
        destAddress = ws.Cells(i, 4).Value
        destFileName = ws.Cells(i, 5).Value
        
        ' エラーメッセージ初期化
        errorMessage = "-"
        
        ' 必須項目の空白チェック
        If processType = "" Or srcAddress = "" Or srcFileName = "" Or destAddress = "" Or destFileName = "" Then
            errorMessage = "空白のセルがあります"
        Else
            ' 移動元のフォルダまたはファイルが存在するかチェック
            If Dir(srcAddress & "\" & srcFileName) = "" Then
                errorMessage = "移動元のフォルダまたはファイルが存在しません"
            Else
                ' 移動先フォルダが存在するかチェック
                If Dir(destAddress, vbDirectory) = "" Then
                    errorMessage = "移動先フォルダが存在しません"
                Else
                    ' 処理内容ごとに処理を分岐
                    Select Case processType
                        Case "移動する（同名を上書きしない）"
                            If Dir(destAddress & "\" & destFileName) <> "" Then
                                errorMessage = "移動先フォルダに同名ファイルが存在しています"
                            Else
                                FileCopy srcAddress & "\" & srcFileName, destAddress & "\" & destFileName
                                Kill srcAddress & "\" & srcFileName
                            End If
                        Case "移動する（同名を上書きする）"
                            FileCopy srcAddress & "\" & srcFileName, destAddress & "\" & destFileName
                            Kill srcAddress & "\" & srcFileName
                        Case "コピーする（同名を上書きしない）"
                            If Dir(destAddress & "\" & destFileName) <> "" Then
                                errorMessage = "コピー先フォルダに同名ファイルが存在しています"
                            Else
                                FileCopy srcAddress & "\" & srcFileName, destAddress & "\" & destFileName
                            End If
                        Case "コピーする（同名を上書きする）"
                            FileCopy srcAddress & "\" & srcFileName, destAddress & "\" & destFileName
                    End Select
                End If
            End If
        End If
        
        ' 処理結果をシートに書き込み
        status = IIf(errorMessage = "-", "完了", "エラー")
        executeTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
        ws.Cells(i, 6).Value = status
        ws.Cells(i, 7).Value = executeTime
        ws.Cells(i, 8).Value = errorMessage
    Next i
End Sub

