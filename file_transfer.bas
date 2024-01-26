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
    
    ' �V�[�g���擾�i�V�[�g���͕K�v�ɉ����ĕύX���Ă��������B�j
    Set ws = ThisWorkbook.Sheets("���C��")
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' �������e���Ƃɏ������s��
    For i = 2 To lastRow
        ' �K�{���ڂ̎擾
        processType = ws.Cells(i, 1).Value
        srcAddress = ws.Cells(i, 2).Value
        srcFileName = ws.Cells(i, 3).Value
        destAddress = ws.Cells(i, 4).Value
        destFileName = ws.Cells(i, 5).Value
        
        ' �G���[���b�Z�[�W������
        errorMessage = "-"
        
        ' �K�{���ڂ̋󔒃`�F�b�N
        If processType = "" Or srcAddress = "" Or srcFileName = "" Or destAddress = "" Or destFileName = "" Then
            errorMessage = "�󔒂̃Z��������܂�"
        Else
            ' �ړ����̃t�H���_�܂��̓t�@�C�������݂��邩�`�F�b�N
            If Dir(srcAddress & "\" & srcFileName) = "" Then
                errorMessage = "�ړ����̃t�H���_�܂��̓t�@�C�������݂��܂���"
            Else
                ' �ړ���t�H���_�����݂��邩�`�F�b�N
                If Dir(destAddress, vbDirectory) = "" Then
                    errorMessage = "�ړ���t�H���_�����݂��܂���"
                Else
                    ' �������e���Ƃɏ����𕪊�
                    Select Case processType
                        Case "�ړ�����i�������㏑�����Ȃ��j"
                            If Dir(destAddress & "\" & destFileName) <> "" Then
                                errorMessage = "�ړ���t�H���_�ɓ����t�@�C�������݂��Ă��܂�"
                            Else
                                FileCopy srcAddress & "\" & srcFileName, destAddress & "\" & destFileName
                                Kill srcAddress & "\" & srcFileName
                            End If
                        Case "�ړ�����i�������㏑������j"
                            FileCopy srcAddress & "\" & srcFileName, destAddress & "\" & destFileName
                            Kill srcAddress & "\" & srcFileName
                        Case "�R�s�[����i�������㏑�����Ȃ��j"
                            If Dir(destAddress & "\" & destFileName) <> "" Then
                                errorMessage = "�R�s�[��t�H���_�ɓ����t�@�C�������݂��Ă��܂�"
                            Else
                                FileCopy srcAddress & "\" & srcFileName, destAddress & "\" & destFileName
                            End If
                        Case "�R�s�[����i�������㏑������j"
                            FileCopy srcAddress & "\" & srcFileName, destAddress & "\" & destFileName
                    End Select
                End If
            End If
        End If
        
        ' �������ʂ��V�[�g�ɏ�������
        status = IIf(errorMessage = "-", "����", "�G���[")
        executeTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
        ws.Cells(i, 6).Value = status
        ws.Cells(i, 7).Value = executeTime
        ws.Cells(i, 8).Value = errorMessage
    Next i
End Sub

