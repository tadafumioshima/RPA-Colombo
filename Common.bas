Attribute VB_Name = "PublicModule"
'----------------------------------------------------------------------------------
'�萔�錾
'----------------------------------------------------------------------------------

Public Const shtItemCheckSheet = "���ڃ`�F�b�N"            '���ڃ`�F�b�N�V�[�g�̃V�[�g��
Public Const shtDataListSheet = "�f�[�^�ꗗ"               '�f�[�^�ꗗ�V�[�g�̃V�[�g��
Public Const shtDataListSheet_2 = "�f�[�^�ꗗ2"            '�f�[�^�ꗗ�V�[�g�̃V�[�g��
Public Const shtFilePassSheet = "�t�@�C���p�X"             '�t�@�C���p�X�V�[�g�̃V�[�g��

Public Const cntBackSlash = "\"                            '\�L�� �t�H���_�̋�؂�
Public Const cntIniFileName = "Sh_Common.ini"              'INI�t�@�C����
Public Const cntConfigPath = "config\"                     'Config�t�H���_��
Public Const cntChrCode_a = &H61                           '"a"��Ascii�R�[�h
Public Const cntNoSettingChr = "-"                         '���ڐݒ��ݒ�L�����N�^

Public Const rowFirstSheetRow = 1                          '���ڃ`�F�b�N�V�[�g�擪�s
Public Const colSheetName = 2                              '�V�[�g����ʒu
Public Const colItemNo = 4                                 '���ڐ���ʒu

Public Const colListInitVal = 8                            '�ꗗ���s��

'---------------------------------------------------------------------------------------
Public Const errNoError = 0                                '�G���[�Ȃ�
Public Const errPeculiarError = 20                         '�c�[���ŗL�G���[
Public Const errWideError = 21                             '�S�p������G���[
Public Const errNarrowError = 30                           '���p������G���[
Public Const errNumericError = 40                          '���l�`���G���[
Public Const errLengthError = 50                           '�������߃G���[
Public Const errLengthLessError = 51                       '�����ߏ��G���[
Public Const errDBConnectionhError = 60                    'DB�ڑ��G���[
Public Const errDBUpdateError = 61                         'DB�X�V�G���[
Public Const errDBInsertError = 67                         'DB�}���G���[
Public Const errDBDeleteError = 68                         'DB�폜�G���[
Public Const errCSVFileOpenError = 62                      'CSV�t�@�C���ǂݍ��݃G���[
Public Const errCSVFileOutputError = 63                    'CSV�t�@�C���o�̓G���[
Public Const errCSV_FileDeleteError = 64                   'CSV�t�@�C���폜�G���[
Public Const errText_FileOpenError = 65                    '�e�L�X�g�t�@�C���ǂݍ��݃G���[
Public Const errText_FileOutputError = 66                  '�e�L�X�g�t�@�C���o�̓G���[
Public Const errRecordSetError = 70                        '���R�[�h�Z�b�g�G���[
Public Const errRecordSetSetError = 71                     '���R�[�h�Z�b�g�Z�b�g�G���[
Public Const errNoDataError = 80                           '�f�[�^�Ȃ�
Public Const errIllegalInputError = 81                     '�s������
Public Const errNoInputError = 82                          '������
Public Const errNoOutputError = 83                         '�o�̓f�[�^�G���[
Public Const errNoInputDataError = 84                      '�f�[�^�ǂݍ��݃G���[
Public Const errIniFileReadError = 95                      'INI�t�@�C���ǂݍ��݃G���[
Public Const errUnknownError = 99                          '�s���G���[


'---------------------------------------------------------------------------------------


'----------------------------------------------------------------------------------
'���ʕϐ��錾
'----------------------------------------------------------------------------------

Public strDBConnection As String                           'DB�ڑ�������i�[�ϐ�
Public strCsvOpen As String                                'CSV�t�@�C���p�X�i�[�ϐ�
Public strTextOpen As String                               '�e�L�X�g�t�@�C���p�X�i�[�ϐ�
Public strExecProc As String                               '���s�v���V�[�W����

Public shtItemWriteSheet As String                         '���o�f�[�^���o���p�̃V�[�g

Public strErrItem As String                                '�G���[����
Public strErrTenCD As String                               '�G���[�X��
Public strErrCIFCD As String                               '�G���[�ڋq�ԍ�
Public strErrPosition As String
Public strErrLength As String

Public intItemRowPosition As Long                          '���ڍs�ʒu
Public intItemColPosition As Long                          '���ڗ�ʒu
Public FormRow As Long

Public strErrTitle As String
Public strErrMsg As String

Public intMaxRow As Long

Public Int_FF As Long
Public fs

'�O����`�錾
'ini�t�@�C���擾
#If Win64 Then
Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#End If

'�`�F�b�N����
Public Function Input_Check(ByVal strItemCheckSheet As String) As Integer

    Dim process_code As Integer
    Dim strSheetName As String                                        '�V�[�g��
    Dim i As Long, j As Long, k As Long, l As Long, m As Long
    Dim intCount As Long, intItemSu As Long
    Dim intItemAttribute As Long                                   '���ڑ���
    Dim intDigits As Long
    Dim intItemCol As Long, intItemRow As Long

    strExecProc = "���̓`�F�b�N"

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = errUnknownError
    Input_Check = process_code

    '�ΏۃV�[�g�擾
    strSheetName = GetMainSheetName(strItemCheckSheet)

    '���ڐ��擾
    intItemSu = Sheets(strItemCheckSheet).Cells(1, 4).Value

    '�ꗗ�����擾
    intCount = Sheets(strItemCheckSheet).Cells(1, 8).Value

    '���̓`�F�b�N
    If intCount > 0 Then
        For i = 1 To intItemSu
            DoEvents
            '���ڃ`�F�b�N�V�[�g����Y�����ڂ̑����擾
            intItemAttribute = Sheets(strItemCheckSheet).Cells(i + 2, 6).Value
            intItemCol = Sheets(strItemCheckSheet).Cells(i + 2, 5).Value
            intItemRow = Sheets(strItemCheckSheet).Cells(i + 2, 4).Value

            If intItemAttribute = 0 Then
            Else
                Select Case intItemAttribute
                    '���l�`�F�b�N
                    Case 4
                        For j = intItemRow To intCount + Sheets(strItemCheckSheet).Cells(i + 2, 4).Value - 1
                            If IsNumeric(Sheets(strSheetName).Cells(j, intItemCol).Value) = False _
                            And Sheets(strSheetName).Cells(j, intItemCol).Value <> "" Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value

                                intItemRowPosition = j
                                intItemColPosition = intItemCol
                                '�߂�l�ɐ��l�G���[�Z�b�g
                                process_code = errNumericError
                                GoTo error_rtn
                            End If
                        Next j
                    '�S�p�`�F�b�N
                    Case 2
                        For l = intItemRow To intCount + Sheets(strItemCheckSheet).Cells(i + 2, 4).Value - 1
                            If Sheets(strSheetName).Cells(l, intItemCol).Value <> StrConv(Sheets(strSheetName).Cells(l, intItemCol).Value, vbWide) Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value
                                intItemRowPosition = l
                                intItemColPosition = intItemCol
                                '�߂�l�ɐ��l�G���[�Z�b�g
                                process_code = errWideError
                                GoTo error_rtn
                            End If
                        Next l
                    '���p�`�F�b�N
                    Case 3
                        For m = intItemRow To intCount + Sheets(strItemCheckSheet).Cells(i + 2, 4).Value - 1
                            If Len(Sheets(strSheetName).Cells(m, intItemCol).Value) <> LenB(StrConv(Sheets(strSheetName).Cells(m, intItemCol).Value, vbFromUnicode)) Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value
                                intItemRowPosition = m
                                intItemColPosition = intItemCol
                                '�߂�l�ɐ��l�G���[�Z�b�g
                                process_code = errNarrowError
                                GoTo error_rtn
                            End If
                        Next m
                    Case Else
                End Select
                '�����`�F�b�N
                intDigits = Sheets(strItemCheckSheet).Cells(i + 2, 7).Value
                If intDigits > 0 And IsNumeric(intDigits) = True Then
                    For k = intItemRow To intCount + Sheets(strItemCheckSheet).Cells(i + 2, 4).Value - 1
                        If intItemAttribute = "2" Then
                            If LenB(StrConv(Sheets(strSheetName).Cells(k, intItemCol).Value, vbFromUnicode)) > intDigits * 2 Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value
                                strErrLength = intDigits
                                intItemRowPosition = k
                                intItemColPosition = intItemCol
                                '�߂�l�Ɍ����G���[�Z�b�g
                                process_code = errLengthError
                                GoTo error_rtn
                            End If
                        Else
                            If LenB(StrConv(Format(Sheets(strSheetName).Cells(k, intItemCol).Value, "0"), vbFromUnicode)) > intDigits Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value
                                strErrLength = intDigits
                                intItemRowPosition = k
                                intItemColPosition = intItemCol
                                '�߂�l�Ɍ����G���[�Z�b�g
                                process_code = errLengthError
                                GoTo error_rtn
                            End If
                        End If
                    Next k
                End If
            End If
        Next i
    End If

'����I�����[�`��
legal_end_rtn:

    '����I��
    Input_Check = errNoError
    Exit Function

'�G���[���[�`��
error_rtn:

    '�w�i�F�̕ύX
    Sheets(strSheetName).Cells(intItemRowPosition, intItemColPosition).Interior.Color = vbRed

    '�J�[�\������V�[�g�G���[�����ꏊ�ɐݒ�
    Call Cursor_Set(strSheetName, intItemColPosition, intItemRowPosition)

    '�߂�l�ݒ�
    Input_Check = process_code
    Exit Function

End Function

'�u�b�N�̂���t�H���_�̐e�t�H���_�擾
Public Function Get_Parent_Folder() As String
    '�G���[�������G���[���[�`����
    On Error GoTo error_rtn

    Dim i As Integer
    Dim strPath As String

    '�u�b�N�̃t�H���_�ݒ�
    strPath = ThisWorkbook.Path

    '�e�t�H���_�擾
    If strPath <> "C:\" Then
        For i = Len(strPath) To 1 Step -1
            DoEvents
            If Mid(strPath, i, 1) = cntBackSlash Then
                strPath = Left(strPath, i)
                Exit For
            End If
        Next i
    End If

lega_rtn:
    '�߂�l��ݒ�
    Get_Parent_Folder = strPath
    Exit Function

error_rtn:
    Get_Parent_Folder = "err"

End Function

'INI�t�@�C���ǂݍ���
Public Function Read_InitFile(ByVal strItemCheckSheet As String) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "INI�t�@�C���ǂݍ���"

    '�߂�l��INI�t�@�C���ǂݍ��݃G���[�Z�b�g
    process_code = errIniFileReadError
    Read_InitFile = process_code

    'ini�t�@�C���̊i�[�ꏊ�̓G�N�Z���c�[���̊i�[����Ă���t�H���_�Ɠ����config�t�H���_�Ƃ���B
    Dim i As Integer
    Dim strPath As String

    strPath = Get_Parent_Folder()
        If Right(strPath, 1) <> cntBackSlash Then strPath = strPath & cntBackSlash

        strPath = strPath & cntConfigPath

    'ini�t�@�C�������t���p�X�ō쐬
    strPath = strPath & cntIniFileName

    Dim buf As String * 256
    'DB�ڑ�������擾
    strDBConnection = ""
    '�v���o�C�_�擾
'    i = GetPrivateProfileString("DB", "Provider", vbNullChar, buf, Len(buf), strPath)
'    If i = 0 Then GoTo error_rtn
'    strDBConnection = "Provider=" & Left(buf, InStr(buf, vbNullChar) - 1)
    '�f�[�^�\�[�X�擾
    i = GetPrivateProfileString("DB", "DSN", vbNullChar, buf, Len(buf), strPath)
    If i = 0 Then GoTo error_rtn
    strDBConnection = strDBConnection & "DSN=" & Left(buf, InStr(buf, vbNullChar) - 1)
    '���[�UID�擾
    i = GetPrivateProfileString("DB", "USER ID", vbNullChar, buf, Len(buf), strPath)
    If i = 0 Then GoTo error_rtn
    strDBConnection = strDBConnection & "USER ID=" & Left(buf, InStr(buf, vbNullChar) - 1)
    '�p�X���[�h�擾
    i = GetPrivateProfileString("DB", "PASSWORD", vbNullChar, buf, Len(buf), strPath)
    If i = 0 Then GoTo error_rtn
    strDBConnection = strDBConnection & "PASSWORD=" & Left(buf, InStr(buf, vbNullChar) - 1)

    'CSV�t�@�C���p�X�擾
    strCsvOpen = ""
    '�t�@�C���p�X�擾
    i = GetPrivateProfileString("CSV_" & Sheets(strItemCheckSheet).Cells(1, 2).Value, "FilePass", vbNullChar, buf, Len(buf), strPath)
    If i <> 0 Then
        strCsvOpen = Left(buf, InStr(buf, vbNullChar) - 1)
    End If
    '�t�@�C�����擾
    i = GetPrivateProfileString("CSV_" & Sheets(strItemCheckSheet).Cells(1, 2).Value, "FileName", vbNullChar, buf, Len(buf), strPath)
    If i <> 0 Then
        strCsvOpen = strCsvOpen & Left(buf, InStr(buf, vbNullChar) - 1)
    End If

    '�e�L�X�g�t�@�C���p�X�擾
    strTextOpen = ""
    '�t�@�C���p�X�擾
    i = GetPrivateProfileString("TEXT_" & Sheets(strItemCheckSheet).Cells(1, 2).Value, "FilePass", vbNullChar, buf, Len(buf), strPath)
    If i <> 0 Then
        strTextOpen = Left(buf, InStr(buf, vbNullChar) - 1)
    End If
    '�t�@�C�����擾
    i = GetPrivateProfileString("TEXT_" & Sheets(strItemCheckSheet).Cells(1, 2).Value, "FileName", vbNullChar, buf, Len(buf), strPath)
    If i <> 0 Then
        strTextOpen = strTextOpen & Left(buf, InStr(buf, vbNullChar) - 1)
    End If

'����I�����[�`��
legal_end_rtn:
    '����I��
    Read_InitFile = errNoError
    Exit Function

'�G���[���[�`��
error_rtn:
    Read_InitFile = process_code
    Exit Function

End Function

'DB�ڑ�
Public Function DB_Connection(ByRef db As Object) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "DB�ڑ�"

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = errDBConnectionhError
    DB_Connection = process_code

    'Set db = CreateObject("ADODB.Connection")
        'db.Open strDBConnection

'����I�����[�`��
legal_end_rtn:
    '����I��
    DB_Connection = errNoError
    Exit Function

'�G���[���[�`��
error_rtn:
    DB_Connection = process_code
    Exit Function

End Function

'CSV�t�@�C���Ǎ�
Public Function CSV_FileOpen(ByRef fs, ByRef Int_FF As Long) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer
    
    strExecProc = "CSV�ǂݍ���"

    process_code = errCSVFileOpenError
    CSV_FileOpen = process_code
               
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Int_FF = FreeFile
    
    Open strCsvOpen For Input As #Int_FF
    
'����I�����[�`��
legal_end_rtn:
    '����I��
    CSV_FileOpen = errNoError
    Exit Function

'�G���[���[�`��
error_rtn:
    CSV_FileOpen = process_code
    Exit Function

End Function

'CSV�t�@�C���o��
Public Function CSV_FileOutput(ByRef fs, ByRef Int_FF As Long, ByVal strDataListSheet As String, _
                            ByVal strItemCheckSheet As String, ByVal intItemNo As Long, ByVal intLastItemNo As Long) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer
    Dim intOutputItemNo As Long
    Dim strOutputItemList() As String                                '�o�͍��ڈꎞ�i�[�ꏊ
    Dim strData As String
    Dim i As Long, j As Long, k As Long, Count As Long

    Count = 1

    strExecProc = "CSV�o��"

    process_code = errCSVFileOutputError
    CSV_FileOutput = process_code

    '�����ڐ��擾
    intOutputItemNo = intItemNo * intLastItemNo

'    '�o�͂���f�[�^�����݂��Ȃ��ꍇ
'    If intLastItemNo < 1 Then
'        process_code = errNoOutputError
'        GoTo error_rtn
'    End If

    Set fs = CreateObject("Scripting.FileSystemObject")

    Int_FF = FreeFile

    Open strCsvOpen For Output As #Int_FF

    If intLastItemNo > 0 Then
        ReDim strOutputItemList(intOutputItemNo)

        '�o�͓��e���擾
        For i = 1 To intLastItemNo
            DoEvents
            For j = 1 To intItemNo
                strOutputItemList(Count) = Sheets(strDataListSheet).Cells(i, j).Value
                Count = Count + 1
            Next j
        Next i

        For k = 1 To intOutputItemNo
            DoEvents
            If k = 1 Then
            ElseIf k Mod intItemNo = 1 Then
                strData = strData & vbCrLf
            Else
                strData = strData & ","
            End If
            strData = strData & strOutputItemList(k)
        Next k

        '�f�[�^�̃t�@�C���ւ̏����o��
        If Len(strData) > 0 Then
            Print #Int_FF, strData
        End If
    End If

    Close #Int_FF

'����I�����[�`��
legal_end_rtn:
    '����I��
    CSV_FileOutput = errNoError
    Exit Function

'�G���[���[�`��
error_rtn:
    CSV_FileOutput = process_code
    Exit Function
    
    Close #Int_FF

End Function

'CSV�t�@�C���폜
Public Function CSV_FileDelete(ByRef fs) As Integer

    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer
    
    strExecProc = "CSV�폜"

    process_code = errCSV_FileDeleteError
    CSV_FileDelete = process_code
        
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Call fs.deletefile(strCsvOpen)
        
'����I�����[�`��
legal_end_rtn:
    '����I��
    CSV_FileDelete = errNoError
    Exit Function

'�G���[���[�`��
error_rtn:
    CSV_FileDelete = process_code
    Exit Function

End Function

'�e�L�X�g�t�@�C���Ǎ�
Public Function Text_FileOpen(ByRef fs, ByRef Int_FF As Long) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer
    
    strExecProc = "�e�L�X�g�t�@�C���ǂݍ���"

    process_code = errText_FileOpenError
    Text_FileOpen = process_code
               
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Int_FF = FreeFile
    
    Open strTextOpen For Input As #Int_FF
    
'����I�����[�`��
legal_end_rtn:
    '����I��
    Text_FileOpen = errNoError
    Exit Function

'�G���[���[�`��
error_rtn:
    Text_FileOpen = process_code
    Exit Function

End Function

'�e�L�X�g�t�@�C���o��
Public Function Text_FileOutput(ByRef fs, ByRef Int_FF As Long, ByVal strDataListSheet As String, _
                            ByVal intItemNo As Long, ByVal intLastItemNo As Long) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer
    Dim intOutputItemNo As Long
    Dim strOutputItemList() As String                                '�o�͍��ڈꎞ�i�[�ꏊ
    Dim strData As String
    Dim i As Long, j As Long, k As Long, Count As Long

    strExecProc = "�e�L�X�g�t�@�C���o��"

    Count = 1

    process_code = errText_FileOutputError
    Text_FileOutput = process_code

    '�����ڐ��擾
    intOutputItemNo = intItemNo * intLastItemNo

    '�o�͂���f�[�^�����݂��Ȃ��ꍇ
    'If intLastItemNo < 1 Then
    '    process_code = errNoOutputError
    '    GoTo error_rtn
    'End If

    Set fs = CreateObject("Scripting.FileSystemObject")

    Int_FF = FreeFile

    Open strTextOpen For Output As #Int_FF

    ReDim strOutputItemList(intOutputItemNo)

    '�o�͓��e���擾
    For i = 1 To intLastItemNo
        DoEvents
        For j = 1 To intItemNo
            strOutputItemList(Count) = Sheets(strDataListSheet).Cells(i, j).Value
            Count = Count + 1
        Next j
    Next i

    For k = 1 To intOutputItemNo
        DoEvents
        If k = 1 Then
        ElseIf k Mod intItemNo = 1 Then
            strData = strData & vbCrLf
        Else
            strData = strData & vbCrLf
        End If
        strData = strData & strOutputItemList(k)
    Next k

    '�f�[�^�̃t�@�C���ւ̏����o��
    If Len(strData) > 0 Then
        Print #Int_FF, strData
    End If
 
    Close #Int_FF

'����I�����[�`��
legal_end_rtn:
    '����I��
    Text_FileOutput = errNoError
    Exit Function
'�G���[���[�`��
error_rtn:
    Text_FileOutput = process_code
    Exit Function
    Close #Int_FF
End Function

'���R�[�h�Z�b�g�擾
Public Function Get_RS(ByRef db As Object, ByRef rs As Object, ByVal strSQL As String) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "�f�[�^�擾"

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = errRecordSetError
    Get_RS = process_code

    Set rs = CreateObject("ADODB.RecordSet")

    rs.Open strSQL, db

    If rs.EOF = True Then
        '�߂�l�Ƀf�[�^�Ȃ��Z�b�g
        Get_RS = errNoDataError
        Exit Function
    End If

'����I�����[�`��
legal_end_rtn:
    '����I��
    Get_RS = errNoError
    Exit Function
'�G���[���[�`��
error_rtn:
    Get_RS = process_code
    Exit Function
End Function

'���R�[�h�X�V
Public Function Get_UP(ByRef db As Object, ByRef rs As Object, ByVal strSQL As String, No As Integer) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "�f�[�^�X�V"

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = No
    Get_UP = process_code

    Set rs = CreateObject("ADODB.RecordSet")
    Set rs = db.Execute(strSQL, i)

    If i = 0 Then
        Exit Function
    End If

'����I�����[�`��
legal_end_rtn:
    '����I��
    Get_UP = errNoError
    Exit Function
'�G���[���[�`��
error_rtn:
    Get_UP = process_code
    Exit Function
End Function
'���R�[�h�폜
Public Function Get_DL(ByRef db As Object, ByVal strSQL As String, No As Integer) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "�f�[�^�폜"

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = No
    Get_DL = process_code

    db.Execute strSQL

'����I�����[�`��
legal_end_rtn:
    '����I��
    Get_DL = errNoError
    Exit Function
'�G���[���[�`��
error_rtn:
    Get_DL = process_code
    Exit Function
End Function
'���R�[�h�ǉ�
Public Function Get_AD(ByRef db As Object, ByVal strSQL As String, No As Integer) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "�f�[�^�ǉ�"

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = No
    Get_AD = process_code

    db.Execute strSQL

'����I�����[�`��
legal_end_rtn:
    '����I��
    Get_AD = errNoError
    Exit Function
'�G���[���[�`��
error_rtn:
    Get_AD = process_code
    Exit Function
End Function
'���R�[�h�Z�b�g�̓��e���V�[�g�ɐݒ�
Public Function Set_RS_To_Sheet(ByRef rs As Object, ByVal shtItemWriteSheet As String) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim process_code As Integer
    
    Dim i As Long
    Dim j As Long

    strExecProc = "�擾�f�[�^�W�J"

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = errRecordSetSetError
    Set_RS_To_Sheet = process_code

    '���R�[�h�Z�b�g�ǂݍ���
    If rs.EOF = False Then
        rs.MoveFirst
    End If

    j = 1

    Do Until rs.EOF = True
        DoEvents
        For i = 1 To rs.Fields.Count
            Sheets(shtItemWriteSheet).Cells(j, i).Value = rs(i - 1)
        Next i
        j = j + 1
        rs.MoveNext
    Loop

'����I�����[�`��
legal_end_rtn:

    '����I��
    Set_RS_To_Sheet = errNoError
    Exit Function

'�G���[���[�`��
error_rtn:
    
    Set_RS_To_Sheet = process_code
    Exit Function

End Function

'�V�[�g������
Public Function Initialize_Sheet(ByVal strItemCheckSheet As String, ByVal strDataListSheet As String) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn
    
    Dim process_code As Integer
    
    Dim strSheetName As String
    Dim intItemNo As Long
    Dim intListInitVal As Long

    Dim i As Long, j As Long
    Dim lngLastRow As Long
    Dim intLockRow As Long, intLockCol As Long
    Dim ws1 As Worksheet, ws2 As Worksheet

    strExecProc = "�V�[�g������"

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = errUnknownError
    Initialize_Sheet = process_code

    '�ΏۃV�[�g�擾
    strSheetName = GetMainSheetName(strItemCheckSheet)

    Set ws1 = Sheets(strSheetName)
    Set ws2 = Sheets(strItemCheckSheet)

    '���ڐ��擾
    intItemNo = ws2.Cells(rowFirstSheetRow, colItemNo).Value
    
    '�ꗗ���s���擾
    intListInitVal = ws2.Cells(rowFirstSheetRow, colListInitVal).Value

    If intListInitVal <> 0 Then
        For i = 3 To intItemNo + 2
            DoEvents
            If ws2.Cells(i, 2).Value = 1 Then
                Select Case ws2.Cells(i, 3).Value
                    '�P���f�[�^�̏ꍇ
                    Case 0
                        With ws1.Cells(ws2.Cells(i, 4).Value, ws2.Cells(i, 5).Value)
                            .ClearContents
                            .Interior.ColorIndex = 0
                            .Borders.LineStyle = xlLineStyleNone
                        End With
                    '�ꗗ�f�[�^�̏ꍇ
                    Case 1
                        With ws1.Range(ws1.Cells(ws2.Cells(i, 4).Value, ws2.Cells(i, 5).Value), ws1.Cells(intListInitVal + ws2.Cells(i, 4).Value - 1, ws2.Cells(i, 5).Value))
                            .ClearContents
                            .Interior.ColorIndex = 0
                            .Borders.LineStyle = xlLineStyleNone
                        End With
                    Case Else
                End Select
            End If
        Next i
    End If
    
    '���C���V�[�g�́u�Z���̃��b�N�v��L���ɂ���
    ws1.Cells.Locked = True
    
    Sheets(strDataListSheet).Cells.Clear
        
    ws2.Cells(rowFirstSheetRow, colListInitVal).Value = 0
    
    Set ws1 = Nothing
    Set ws2 = Nothing
        
'����I�����[�`��
legal_end_rtn:
    '����I��
    Initialize_Sheet = errNoError
    Exit Function
    
'�G���[���[�`��
error_rtn:
    Initialize_Sheet = process_code
    Exit Function
    
End Function

'�w�i�F������
Public Function Initialize_Paint_Sheet(ByVal strItemCheckSheet As String) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn
    
    Dim process_code As Integer
    
    Dim strSheetName As String
    Dim intItemNo As Long
    Dim intListInitVal As Long
                        
    Dim i As Long, j As Long
    Dim lngLastRow As Long
    
    strExecProc = "�V�[�g������(�w�i�F)"

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = errUnknownError
    Initialize_Paint_Sheet = process_code
            
    '�ΏۃV�[�g�擾
    strSheetName = GetMainSheetName(strItemCheckSheet)
    '���ڐ��擾
    intItemNo = Sheets(strItemCheckSheet).Cells(rowFirstSheetRow, colItemNo).Value
    
    '�ꗗ���s���擾
    intListInitVal = Sheets(strItemCheckSheet).Cells(rowFirstSheetRow, colListInitVal).Value
        
    If intListInitVal <> 0 Then
        For i = 3 To intItemNo + 2
            DoEvents
            If Sheets(strItemCheckSheet).Cells(i, 2).Value = 1 Then
                Select Case Sheets(strItemCheckSheet).Cells(i, 3).Value
                    '�P���f�[�^�̏ꍇ
                    Case 0
                        Sheets(strSheetName).Cells(Sheets(strItemCheckSheet).Cells(i, 4).Value, Sheets(strItemCheckSheet).Cells(i, 5).Value).Interior.ColorIndex = 0
                    '�ꗗ�f�[�^�̏ꍇ
                    Case 1
                        For j = Sheets(strItemCheckSheet).Cells(i, 4).Value To intListInitVal + Sheets(strItemCheckSheet).Cells(i, 4).Value - 1
                            Sheets(strSheetName).Cells(j, Sheets(strItemCheckSheet).Cells(i, 5).Value).Interior.ColorIndex = 0
                        Next j
                    Case Else
                End Select
            End If
        Next i
    End If
    
'����I�����[�`��
legal_end_rtn:
    '����I��
    Initialize_Paint_Sheet = errNoError
    Exit Function
    
'�G���[���[�`��
error_rtn:
    Initialize_Paint_Sheet = process_code
    Exit Function
    
End Function

'��V�[�g���擾
Public Function GetMainSheetName(ByVal strItemCheckSheet As String) As String
    GetMainSheetName = Sheets(strItemCheckSheet).Cells(rowFirstSheetRow, colSheetName).Value
End Function

'�J�[�\���ݒ�
Public Sub Cursor_Set(ByVal strSheetName As String, ByVal Y As Long, ByVal X As Long)
    Dim strPreChr As String
    
    strPreChr = ""
    If Y > 26 Then
        strPreChr = Chr(cntChrCode_a + (Y \ 26) - 1 + IIf(Y Mod 26 = 0, -1, 0))
        Y = Y Mod 26
    End If
    
    '�����̈ʒu�ɃJ�[�\����ݒ�
    Application.CutCopyMode = False
    ActiveCell.Copy
    Worksheets(strSheetName).Range(strPreChr & Chr(cntChrCode_a + IIf(Y - 1 > 0, Y - 1, 0)) & X).PasteSpecial xlPasteComments
    Application.CutCopyMode = True
End Sub

'�ꗗ�ŏI�s�擾
Public Function Items_Last_Row(ByVal strItemCheckSheet As String) As Integer
    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim intAllNo As Long, intAllNoCol As Long, intAllNoRow As Long
    Dim intMaxNo As Long, intNoRow As Long, intFastCol As Long, intLastCol As Long
    Dim Items_Col()
    Dim i As Long

    strExecProc = "�ꗗ�ŏI�s�擾"
    
    '�߂�l�ɕs���G���[�Z�b�g
    process_code = errUnknownError
    Items_Last_Row = process_code

    '�ΏۃV�[�g�擾
    strSheetName = GetMainSheetName(strItemCheckSheet)
        
    '�������ʒu�擾
    intAllNo = Sheets(strItemCheckSheet).Cells.Find("������", LookAt:=xlWhole).Row
    intAllNoRow = Sheets(strItemCheckSheet).Cells(intAllNo, 4).Value
    intAllNoCol = Sheets(strItemCheckSheet).Cells(intAllNo, 5).Value

    intMaxNo = 0
    intNoRow = Sheets(strItemCheckSheet).Cells.Find("����", LookAt:=xlWhole).Row
    intFastCol = Sheets(strItemCheckSheet).Cells(intNoRow, 5).Value
    intLastCol = Sheets(strItemCheckSheet).Cells(Sheets(strItemCheckSheet).Cells(1, 4).Value + 2, 5).Value
    
    '�ŏI�s�擾
    ReDim Items_Col(intLastCol)
    Do
        DoEvents
        Items_Col(intFastCol) = Sheets(strSheetName).Cells(Rows.Count, intFastCol).End(xlUp).Row
        intFastCol = intFastCol + 1
        If intFastCol > intLastCol Then
            Exit Do
        End If
    Loop

    For i = 0 To UBound(Items_Col)
        DoEvents
        If Items_Col(i) > intMaxNo Then
            intMaxNo = Items_Col(i)
        End If
    Next i
    intMaxNo = intMaxNo - (Sheets(strItemCheckSheet).Cells(intNoRow, 4).Value - 1)

    '�ŏI�s�����ڂ֔��f
    If intMaxNo > 0 Then
        Sheets(strItemCheckSheet).Cells(1, 8).Value = intMaxNo
        Sheets(strSheetName).Cells(intAllNoRow, intAllNoCol).Value = Format(intMaxNo, "###,###,###,###") & "��"
    Else
        Sheets(strItemCheckSheet).Cells(1, 8).Value = 0
    End If

'����I���� �[�`��
legal_end_rtn:
    '����I��
    Items_Last_Row = errNoError
    Exit Function

'�G���[���[�`��
error_rtn:
    Items_Last_Row = process_code
    Exit Function
End Function

'�`�F�b�N�{�b�N�X�폜
Public Function CheckBox_Delete(ByVal strItemCheckSheet As String) As Integer

    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    Dim strSheetName As String
    Dim obj As Object

    '�߂�l�ɕs���G���[�Z�b�g
    process_code = errUnknownError
    CheckBox_Delete = process_code

    '�ΏۃV�[�g�擾
    strSheetName = GetMainSheetName(strItemCheckSheet)

    For Each obj In Sheets(strSheetName).OLEObjects
        DoEvents
        If obj.Name Like "CheckBox*" Then
            With Sheets(strSheetName).Shapes(obj.Name)
                .Select
                .Delete
            End With
        End If
    Next

'����I���� �[�`��
legal_end_rtn:
    '����I��
    CheckBox_Delete = errNoError
    Exit Function
'�G���[���[�`��
error_rtn:
    CheckBox_Delete = process_code
    Exit Function
End Function

'���엚������
Public Function OpeLog_Write(ByVal strToolNo As String, ByVal strButtonNo As String) As Integer

    '�G���[�����������ꍇ�A�G���[���[�`����
    On Error GoTo error_rtn

    '�������t�@�C���ǂݍ���
    process_code = Read_InitFile("")

    Dim strUserId As String
    strUserId = Workbooks("�}�X�^�����e�i���X���j���[.xlsm").Sheets("�}�X�^�����e�i���X���j���[").Cells(2, 12).Value

    'DB�I�[�v��
    Dim db As Object
    process_code = DB_Connection(db)
    db.BeginTrans

    '���R�[�h�Z�b�g�I�u�W�F�N�g�̐錾
    Dim rs As Object

    '���엚������
    strSQL = ""
    strSQL = " INSERT INTO TT9901 VALUES ("
    strSQL = strSQL & " TO_CHAR(sysdate,'yyyymmddhh24miss')"
    strSQL = strSQL & ",'999'"
    strSQL = strSQL & ",'" & strUserId & "'"
    strSQL = strSQL & ",'" & strToolNo & "'"
    strSQL = strSQL & ",'" & strButtonNo & "'"
    strSQL = strSQL & ",EMPTY_BLOB())"

    Dim No As Integer
    No = errDBInsertError
    process_code = Get_UP(db, rs, strSQL, No)

    db.CommitTrans

    '���R�[�h�Z�b�g�N���[�Y
    rs.Close

    'DB�N���[�Y
    db.Close

'����I���� �[�`��
legal_end_rtn:
    OpeLog_Write = errNoError
    Exit Function
'�G���[���[�`��
error_rtn:
    OpeLog_Write = process_code
    Exit Function
End Function

'�G���[���b�Z�[�W
Public Sub Display_MSG(ByVal errCode As Integer)
    Dim strTTL As String
    Dim strMSG As String

    Select Case errCode
        Case errNoError
            strTTL = "�����I��"
            strMSG = "�����͐���ɏI�����܂���"
        Case errPeculiarError
            strTTL = strErrTitle
            strMSG = strErrMsg
        Case errWideError
            strTTL = "�S�p�����ȊO�̕s������"
            strMSG = strErrItem & "�ɑS�p�����ȊO�̒l�����͂���Ă��܂�"
        Case errNarrowError
            strTTL = "���p�����ȊO�̕s������"
            strMSG = strErrItem & "�ɔ��p�����ȊO�̒l�����͂���Ă��܂�"
        Case errNumericError
            strTTL = "���l�ȊO�̕s������"
            strMSG = strErrItem & "�����l�ł͂���܂���"
        Case errLengthError
            strTTL = "����������"
            strMSG = strErrItem & "�̌���������𒴂��Ă��܂�(����F" & strErrLength & "��)"
        Case errLengthLessError
            strTTL = "�������ߏ�"
            strMSG = strErrItem & "�̌������K��ɖ����Ȃ��l�ł�(�K��F" & strErrLength & "��)"
        Case errDBConnectionhError
            strTTL = "�f�[�^�x�[�X�ڑ��G���["
            strMSG = "�f�[�^�x�[�X�ւ̐ڑ��Ɏ��s���܂���"
        Case errDBUpdateError
            strTTL = "�f�[�^�x�[�X�X�V�G���["
            strMSG = strErrMsg
        Case errDBInsertError
            strTTL = "�f�[�^�x�[�X�ǉ��G���["
            strMSG = strErrMsg
        Case errDBDeleteError
            strTTL = "�f�[�^�x�[�X�폜�G���["
            strMSG = strErrMsg
        Case errRecordSetError
            strTTL = "�f�[�^���o�G���["
            strMSG = "�f�[�^���o�Ɏ��s���܂���"
        Case errRecordSetSetError
            strTTL = "�f�[�^�쐬�G���["
            strMSG = "�f�[�^�͒��o���܂������A�V�[�g�ւ̓W�J�Ɏ��s���܂���"
        Case errNoDataError
            strTTL = "�f�[�^����"
            strMSG = "�Y������f�[�^�͑��݂��܂���"
        Case errIllegalInputError
            strTTL = "�s�����̓G���["
            strMSG = strExecProc & "�̏������ɁA�s���ȓ��͂����݂��邽�߁A�����𒆒f���܂���"
        Case errNoInputError
            strTTL = "�����̓G���["
            strMSG = strExecProc & "�̏������ɁA�����͍��ڂ����o���ꂽ���߁A" & Chr(&HD) & "�����𒆒f���܂���"
        Case errNoOutputError
            strTTL = "�o�̓G���["
            strMSG = "�o�͂���f�[�^�����݂��܂���"
        Case errIniFileReadError
            strTTL = "INI�t�@�C���ǂݍ��݃G���["
            strMSG = "INI�t�@�C���̓ǂݍ��݂Ɏ��s���܂���"
        Case errUnknownError
            strTTL = "�ُ�I��"
            strMSG = strExecProc & "�̏������Ɉُ�I�����܂���"
        Case errCSVFileOpenError
            strTTL = "CSV�ǂݍ��݃G���["
            strMSG = "CSV�t�@�C���̓ǂݍ��݂Ɏ��s���܂���"
        Case errCSVFileOutputError
            strTTL = "CSV�o�̓G���["
            strMSG = "CSV�t�@�C���̏o�͂Ɏ��s���܂���"
        Case errCSV_FileDeleteError
            strTTL = "CSV�폜�G���["
            strMSG = "CSV�t�@�C���̍폜�Ɏ��s���܂���"
        Case errText_FileOpenError
            strTTL = "�e�L�X�g�t�@�C���Ǎ��G���["
            strMSG = "�e�L�X�g�t�@�C���̓Ǎ��Ɏ��s���܂���"
        Case errText_FileOutputError
            strTTL = "�e�L�X�g�t�@�C���o�̓G���["
            strMSG = "�e�L�X�g�t�@�C���̏o�͂Ɏ��s���܂���"
        Case errNoInputDataError
            strTTL = "�Ǎ��G���["
            strMSG = "�t�@�C���Ƀf�[�^�����݂��܂���"
    End Select

    MsgBox strMSG, vbOKOnly, strTTL

End Sub
