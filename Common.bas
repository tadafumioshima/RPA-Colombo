Attribute VB_Name = "PublicModule"
'----------------------------------------------------------------------------------
'定数宣言
'----------------------------------------------------------------------------------

Public Const shtItemCheckSheet = "項目チェック"            '項目チェックシートのシート名
Public Const shtDataListSheet = "データ一覧"               'データ一覧シートのシート名
Public Const shtDataListSheet_2 = "データ一覧2"            'データ一覧シートのシート名
Public Const shtFilePassSheet = "ファイルパス"             'ファイルパスシートのシート名

Public Const cntBackSlash = "\"                            '\記号 フォルダの区切り
Public Const cntIniFileName = "Sh_Common.ini"              'INIファイル名
Public Const cntConfigPath = "config\"                     'Configフォルダ名
Public Const cntChrCode_a = &H61                           '"a"のAsciiコード
Public Const cntNoSettingChr = "-"                         '項目設定非設定キャラクタ

Public Const rowFirstSheetRow = 1                          '項目チェックシート先頭行
Public Const colSheetName = 2                              'シート名列位置
Public Const colItemNo = 4                                 '項目数列位置

Public Const colListInitVal = 8                            '一覧部行数

'---------------------------------------------------------------------------------------
Public Const errNoError = 0                                'エラーなし
Public Const errPeculiarError = 20                         'ツール固有エラー
Public Const errWideError = 21                             '全角文字列エラー
Public Const errNarrowError = 30                           '半角文字列エラー
Public Const errNumericError = 40                          '数値形式エラー
Public Const errLengthError = 50                           '桁数超過エラー
Public Const errLengthLessError = 51                       '桁数過少エラー
Public Const errDBConnectionhError = 60                    'DB接続エラー
Public Const errDBUpdateError = 61                         'DB更新エラー
Public Const errDBInsertError = 67                         'DB挿入エラー
Public Const errDBDeleteError = 68                         'DB削除エラー
Public Const errCSVFileOpenError = 62                      'CSVファイル読み込みエラー
Public Const errCSVFileOutputError = 63                    'CSVファイル出力エラー
Public Const errCSV_FileDeleteError = 64                   'CSVファイル削除エラー
Public Const errText_FileOpenError = 65                    'テキストファイル読み込みエラー
Public Const errText_FileOutputError = 66                  'テキストファイル出力エラー
Public Const errRecordSetError = 70                        'レコードセットエラー
Public Const errRecordSetSetError = 71                     'レコードセットセットエラー
Public Const errNoDataError = 80                           'データなし
Public Const errIllegalInputError = 81                     '不正入力
Public Const errNoInputError = 82                          '未入力
Public Const errNoOutputError = 83                         '出力データエラー
Public Const errNoInputDataError = 84                      'データ読み込みエラー
Public Const errIniFileReadError = 95                      'INIファイル読み込みエラー
Public Const errUnknownError = 99                          '不明エラー


'---------------------------------------------------------------------------------------


'----------------------------------------------------------------------------------
'共通変数宣言
'----------------------------------------------------------------------------------

Public strDBConnection As String                           'DB接続文字列格納変数
Public strCsvOpen As String                                'CSVファイルパス格納変数
Public strTextOpen As String                               'テキストファイルパス格納変数
Public strExecProc As String                               '実行プロシージャ名

Public shtItemWriteSheet As String                         '抽出データ書出し用のシート

Public strErrItem As String                                'エラー項目
Public strErrTenCD As String                               'エラー店番
Public strErrCIFCD As String                               'エラー顧客番号
Public strErrPosition As String
Public strErrLength As String

Public intItemRowPosition As Long                          '項目行位置
Public intItemColPosition As Long                          '項目列位置
Public FormRow As Long

Public strErrTitle As String
Public strErrMsg As String

Public intMaxRow As Long

Public Int_FF As Long
Public fs

'外部定義宣言
'iniファイル取得
#If Win64 Then
Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#End If

'チェック処理
Public Function Input_Check(ByVal strItemCheckSheet As String) As Integer

    Dim process_code As Integer
    Dim strSheetName As String                                        'シート名
    Dim i As Long, j As Long, k As Long, l As Long, m As Long
    Dim intCount As Long, intItemSu As Long
    Dim intItemAttribute As Long                                   '項目属性
    Dim intDigits As Long
    Dim intItemCol As Long, intItemRow As Long

    strExecProc = "入力チェック"

    '戻り値に不明エラーセット
    process_code = errUnknownError
    Input_Check = process_code

    '対象シート取得
    strSheetName = GetMainSheetName(strItemCheckSheet)

    '項目数取得
    intItemSu = Sheets(strItemCheckSheet).Cells(1, 4).Value

    '一覧件数取得
    intCount = Sheets(strItemCheckSheet).Cells(1, 8).Value

    '入力チェック
    If intCount > 0 Then
        For i = 1 To intItemSu
            DoEvents
            '項目チェックシートから該当項目の属性取得
            intItemAttribute = Sheets(strItemCheckSheet).Cells(i + 2, 6).Value
            intItemCol = Sheets(strItemCheckSheet).Cells(i + 2, 5).Value
            intItemRow = Sheets(strItemCheckSheet).Cells(i + 2, 4).Value

            If intItemAttribute = 0 Then
            Else
                Select Case intItemAttribute
                    '数値チェック
                    Case 4
                        For j = intItemRow To intCount + Sheets(strItemCheckSheet).Cells(i + 2, 4).Value - 1
                            If IsNumeric(Sheets(strSheetName).Cells(j, intItemCol).Value) = False _
                            And Sheets(strSheetName).Cells(j, intItemCol).Value <> "" Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value

                                intItemRowPosition = j
                                intItemColPosition = intItemCol
                                '戻り値に数値エラーセット
                                process_code = errNumericError
                                GoTo error_rtn
                            End If
                        Next j
                    '全角チェック
                    Case 2
                        For l = intItemRow To intCount + Sheets(strItemCheckSheet).Cells(i + 2, 4).Value - 1
                            If Sheets(strSheetName).Cells(l, intItemCol).Value <> StrConv(Sheets(strSheetName).Cells(l, intItemCol).Value, vbWide) Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value
                                intItemRowPosition = l
                                intItemColPosition = intItemCol
                                '戻り値に数値エラーセット
                                process_code = errWideError
                                GoTo error_rtn
                            End If
                        Next l
                    '半角チェック
                    Case 3
                        For m = intItemRow To intCount + Sheets(strItemCheckSheet).Cells(i + 2, 4).Value - 1
                            If Len(Sheets(strSheetName).Cells(m, intItemCol).Value) <> LenB(StrConv(Sheets(strSheetName).Cells(m, intItemCol).Value, vbFromUnicode)) Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value
                                intItemRowPosition = m
                                intItemColPosition = intItemCol
                                '戻り値に数値エラーセット
                                process_code = errNarrowError
                                GoTo error_rtn
                            End If
                        Next m
                    Case Else
                End Select
                '桁数チェック
                intDigits = Sheets(strItemCheckSheet).Cells(i + 2, 7).Value
                If intDigits > 0 And IsNumeric(intDigits) = True Then
                    For k = intItemRow To intCount + Sheets(strItemCheckSheet).Cells(i + 2, 4).Value - 1
                        If intItemAttribute = "2" Then
                            If LenB(StrConv(Sheets(strSheetName).Cells(k, intItemCol).Value, vbFromUnicode)) > intDigits * 2 Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value
                                strErrLength = intDigits
                                intItemRowPosition = k
                                intItemColPosition = intItemCol
                                '戻り値に桁数エラーセット
                                process_code = errLengthError
                                GoTo error_rtn
                            End If
                        Else
                            If LenB(StrConv(Format(Sheets(strSheetName).Cells(k, intItemCol).Value, "0"), vbFromUnicode)) > intDigits Then
                                strErrItem = Sheets(strItemCheckSheet).Cells(i + 2, 1).Value
                                strErrLength = intDigits
                                intItemRowPosition = k
                                intItemColPosition = intItemCol
                                '戻り値に桁数エラーセット
                                process_code = errLengthError
                                GoTo error_rtn
                            End If
                        End If
                    Next k
                End If
            End If
        Next i
    End If

'正常終了ルーチン
legal_end_rtn:

    '正常終了
    Input_Check = errNoError
    Exit Function

'エラールーチン
error_rtn:

    '背景色の変更
    Sheets(strSheetName).Cells(intItemRowPosition, intItemColPosition).Interior.Color = vbRed

    'カーソルを主シートエラー発生場所に設定
    Call Cursor_Set(strSheetName, intItemColPosition, intItemRowPosition)

    '戻り値設定
    Input_Check = process_code
    Exit Function

End Function

'ブックのあるフォルダの親フォルダ取得
Public Function Get_Parent_Folder() As String
    'エラー発生時エラールーチンに
    On Error GoTo error_rtn

    Dim i As Integer
    Dim strPath As String

    'ブックのフォルダ設定
    strPath = ThisWorkbook.Path

    '親フォルダ取得
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
    '戻り値を設定
    Get_Parent_Folder = strPath
    Exit Function

error_rtn:
    Get_Parent_Folder = "err"

End Function

'INIファイル読み込み
Public Function Read_InitFile(ByVal strItemCheckSheet As String) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "INIファイル読み込み"

    '戻り値にINIファイル読み込みエラーセット
    process_code = errIniFileReadError
    Read_InitFile = process_code

    'iniファイルの格納場所はエクセルツールの格納されているフォルダと同列のconfigフォルダとする。
    Dim i As Integer
    Dim strPath As String

    strPath = Get_Parent_Folder()
        If Right(strPath, 1) <> cntBackSlash Then strPath = strPath & cntBackSlash

        strPath = strPath & cntConfigPath

    'iniファイル名をフルパスで作成
    strPath = strPath & cntIniFileName

    Dim buf As String * 256
    'DB接続文字列取得
    strDBConnection = ""
    'プロバイダ取得
'    i = GetPrivateProfileString("DB", "Provider", vbNullChar, buf, Len(buf), strPath)
'    If i = 0 Then GoTo error_rtn
'    strDBConnection = "Provider=" & Left(buf, InStr(buf, vbNullChar) - 1)
    'データソース取得
    i = GetPrivateProfileString("DB", "DSN", vbNullChar, buf, Len(buf), strPath)
    If i = 0 Then GoTo error_rtn
    strDBConnection = strDBConnection & "DSN=" & Left(buf, InStr(buf, vbNullChar) - 1)
    'ユーザID取得
    i = GetPrivateProfileString("DB", "USER ID", vbNullChar, buf, Len(buf), strPath)
    If i = 0 Then GoTo error_rtn
    strDBConnection = strDBConnection & "USER ID=" & Left(buf, InStr(buf, vbNullChar) - 1)
    'パスワード取得
    i = GetPrivateProfileString("DB", "PASSWORD", vbNullChar, buf, Len(buf), strPath)
    If i = 0 Then GoTo error_rtn
    strDBConnection = strDBConnection & "PASSWORD=" & Left(buf, InStr(buf, vbNullChar) - 1)

    'CSVファイルパス取得
    strCsvOpen = ""
    'ファイルパス取得
    i = GetPrivateProfileString("CSV_" & Sheets(strItemCheckSheet).Cells(1, 2).Value, "FilePass", vbNullChar, buf, Len(buf), strPath)
    If i <> 0 Then
        strCsvOpen = Left(buf, InStr(buf, vbNullChar) - 1)
    End If
    'ファイル名取得
    i = GetPrivateProfileString("CSV_" & Sheets(strItemCheckSheet).Cells(1, 2).Value, "FileName", vbNullChar, buf, Len(buf), strPath)
    If i <> 0 Then
        strCsvOpen = strCsvOpen & Left(buf, InStr(buf, vbNullChar) - 1)
    End If

    'テキストファイルパス取得
    strTextOpen = ""
    'ファイルパス取得
    i = GetPrivateProfileString("TEXT_" & Sheets(strItemCheckSheet).Cells(1, 2).Value, "FilePass", vbNullChar, buf, Len(buf), strPath)
    If i <> 0 Then
        strTextOpen = Left(buf, InStr(buf, vbNullChar) - 1)
    End If
    'ファイル名取得
    i = GetPrivateProfileString("TEXT_" & Sheets(strItemCheckSheet).Cells(1, 2).Value, "FileName", vbNullChar, buf, Len(buf), strPath)
    If i <> 0 Then
        strTextOpen = strTextOpen & Left(buf, InStr(buf, vbNullChar) - 1)
    End If

'正常終了ルーチン
legal_end_rtn:
    '正常終了
    Read_InitFile = errNoError
    Exit Function

'エラールーチン
error_rtn:
    Read_InitFile = process_code
    Exit Function

End Function

'DB接続
Public Function DB_Connection(ByRef db As Object) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "DB接続"

    '戻り値に不明エラーセット
    process_code = errDBConnectionhError
    DB_Connection = process_code

    'Set db = CreateObject("ADODB.Connection")
        'db.Open strDBConnection

'正常終了ルーチン
legal_end_rtn:
    '正常終了
    DB_Connection = errNoError
    Exit Function

'エラールーチン
error_rtn:
    DB_Connection = process_code
    Exit Function

End Function

'CSVファイル読込
Public Function CSV_FileOpen(ByRef fs, ByRef Int_FF As Long) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer
    
    strExecProc = "CSV読み込み"

    process_code = errCSVFileOpenError
    CSV_FileOpen = process_code
               
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Int_FF = FreeFile
    
    Open strCsvOpen For Input As #Int_FF
    
'正常終了ルーチン
legal_end_rtn:
    '正常終了
    CSV_FileOpen = errNoError
    Exit Function

'エラールーチン
error_rtn:
    CSV_FileOpen = process_code
    Exit Function

End Function

'CSVファイル出力
Public Function CSV_FileOutput(ByRef fs, ByRef Int_FF As Long, ByVal strDataListSheet As String, _
                            ByVal strItemCheckSheet As String, ByVal intItemNo As Long, ByVal intLastItemNo As Long) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer
    Dim intOutputItemNo As Long
    Dim strOutputItemList() As String                                '出力項目一時格納場所
    Dim strData As String
    Dim i As Long, j As Long, k As Long, Count As Long

    Count = 1

    strExecProc = "CSV出力"

    process_code = errCSVFileOutputError
    CSV_FileOutput = process_code

    '総項目数取得
    intOutputItemNo = intItemNo * intLastItemNo

'    '出力するデータが存在しない場合
'    If intLastItemNo < 1 Then
'        process_code = errNoOutputError
'        GoTo error_rtn
'    End If

    Set fs = CreateObject("Scripting.FileSystemObject")

    Int_FF = FreeFile

    Open strCsvOpen For Output As #Int_FF

    If intLastItemNo > 0 Then
        ReDim strOutputItemList(intOutputItemNo)

        '出力内容を取得
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

        'データのファイルへの書き出し
        If Len(strData) > 0 Then
            Print #Int_FF, strData
        End If
    End If

    Close #Int_FF

'正常終了ルーチン
legal_end_rtn:
    '正常終了
    CSV_FileOutput = errNoError
    Exit Function

'エラールーチン
error_rtn:
    CSV_FileOutput = process_code
    Exit Function
    
    Close #Int_FF

End Function

'CSVファイル削除
Public Function CSV_FileDelete(ByRef fs) As Integer

    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer
    
    strExecProc = "CSV削除"

    process_code = errCSV_FileDeleteError
    CSV_FileDelete = process_code
        
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Call fs.deletefile(strCsvOpen)
        
'正常終了ルーチン
legal_end_rtn:
    '正常終了
    CSV_FileDelete = errNoError
    Exit Function

'エラールーチン
error_rtn:
    CSV_FileDelete = process_code
    Exit Function

End Function

'テキストファイル読込
Public Function Text_FileOpen(ByRef fs, ByRef Int_FF As Long) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer
    
    strExecProc = "テキストファイル読み込み"

    process_code = errText_FileOpenError
    Text_FileOpen = process_code
               
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Int_FF = FreeFile
    
    Open strTextOpen For Input As #Int_FF
    
'正常終了ルーチン
legal_end_rtn:
    '正常終了
    Text_FileOpen = errNoError
    Exit Function

'エラールーチン
error_rtn:
    Text_FileOpen = process_code
    Exit Function

End Function

'テキストファイル出力
Public Function Text_FileOutput(ByRef fs, ByRef Int_FF As Long, ByVal strDataListSheet As String, _
                            ByVal intItemNo As Long, ByVal intLastItemNo As Long) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer
    Dim intOutputItemNo As Long
    Dim strOutputItemList() As String                                '出力項目一時格納場所
    Dim strData As String
    Dim i As Long, j As Long, k As Long, Count As Long

    strExecProc = "テキストファイル出力"

    Count = 1

    process_code = errText_FileOutputError
    Text_FileOutput = process_code

    '総項目数取得
    intOutputItemNo = intItemNo * intLastItemNo

    '出力するデータが存在しない場合
    'If intLastItemNo < 1 Then
    '    process_code = errNoOutputError
    '    GoTo error_rtn
    'End If

    Set fs = CreateObject("Scripting.FileSystemObject")

    Int_FF = FreeFile

    Open strTextOpen For Output As #Int_FF

    ReDim strOutputItemList(intOutputItemNo)

    '出力内容を取得
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

    'データのファイルへの書き出し
    If Len(strData) > 0 Then
        Print #Int_FF, strData
    End If
 
    Close #Int_FF

'正常終了ルーチン
legal_end_rtn:
    '正常終了
    Text_FileOutput = errNoError
    Exit Function
'エラールーチン
error_rtn:
    Text_FileOutput = process_code
    Exit Function
    Close #Int_FF
End Function

'レコードセット取得
Public Function Get_RS(ByRef db As Object, ByRef rs As Object, ByVal strSQL As String) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "データ取得"

    '戻り値に不明エラーセット
    process_code = errRecordSetError
    Get_RS = process_code

    Set rs = CreateObject("ADODB.RecordSet")

    rs.Open strSQL, db

    If rs.EOF = True Then
        '戻り値にデータなしセット
        Get_RS = errNoDataError
        Exit Function
    End If

'正常終了ルーチン
legal_end_rtn:
    '正常終了
    Get_RS = errNoError
    Exit Function
'エラールーチン
error_rtn:
    Get_RS = process_code
    Exit Function
End Function

'レコード更新
Public Function Get_UP(ByRef db As Object, ByRef rs As Object, ByVal strSQL As String, No As Integer) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "データ更新"

    '戻り値に不明エラーセット
    process_code = No
    Get_UP = process_code

    Set rs = CreateObject("ADODB.RecordSet")
    Set rs = db.Execute(strSQL, i)

    If i = 0 Then
        Exit Function
    End If

'正常終了ルーチン
legal_end_rtn:
    '正常終了
    Get_UP = errNoError
    Exit Function
'エラールーチン
error_rtn:
    Get_UP = process_code
    Exit Function
End Function
'レコード削除
Public Function Get_DL(ByRef db As Object, ByVal strSQL As String, No As Integer) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "データ削除"

    '戻り値に不明エラーセット
    process_code = No
    Get_DL = process_code

    db.Execute strSQL

'正常終了ルーチン
legal_end_rtn:
    '正常終了
    Get_DL = errNoError
    Exit Function
'エラールーチン
error_rtn:
    Get_DL = process_code
    Exit Function
End Function
'レコード追加
Public Function Get_AD(ByRef db As Object, ByVal strSQL As String, No As Integer) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer

    strExecProc = "データ追加"

    '戻り値に不明エラーセット
    process_code = No
    Get_AD = process_code

    db.Execute strSQL

'正常終了ルーチン
legal_end_rtn:
    '正常終了
    Get_AD = errNoError
    Exit Function
'エラールーチン
error_rtn:
    Get_AD = process_code
    Exit Function
End Function
'レコードセットの内容をシートに設定
Public Function Set_RS_To_Sheet(ByRef rs As Object, ByVal shtItemWriteSheet As String) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim process_code As Integer
    
    Dim i As Long
    Dim j As Long

    strExecProc = "取得データ展開"

    '戻り値に不明エラーセット
    process_code = errRecordSetSetError
    Set_RS_To_Sheet = process_code

    'レコードセット読み込み
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

'正常終了ルーチン
legal_end_rtn:

    '正常終了
    Set_RS_To_Sheet = errNoError
    Exit Function

'エラールーチン
error_rtn:
    
    Set_RS_To_Sheet = process_code
    Exit Function

End Function

'シート初期化
Public Function Initialize_Sheet(ByVal strItemCheckSheet As String, ByVal strDataListSheet As String) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn
    
    Dim process_code As Integer
    
    Dim strSheetName As String
    Dim intItemNo As Long
    Dim intListInitVal As Long

    Dim i As Long, j As Long
    Dim lngLastRow As Long
    Dim intLockRow As Long, intLockCol As Long
    Dim ws1 As Worksheet, ws2 As Worksheet

    strExecProc = "シート初期化"

    '戻り値に不明エラーセット
    process_code = errUnknownError
    Initialize_Sheet = process_code

    '対象シート取得
    strSheetName = GetMainSheetName(strItemCheckSheet)

    Set ws1 = Sheets(strSheetName)
    Set ws2 = Sheets(strItemCheckSheet)

    '項目数取得
    intItemNo = ws2.Cells(rowFirstSheetRow, colItemNo).Value
    
    '一覧部行数取得
    intListInitVal = ws2.Cells(rowFirstSheetRow, colListInitVal).Value

    If intListInitVal <> 0 Then
        For i = 3 To intItemNo + 2
            DoEvents
            If ws2.Cells(i, 2).Value = 1 Then
                Select Case ws2.Cells(i, 3).Value
                    '単項データの場合
                    Case 0
                        With ws1.Cells(ws2.Cells(i, 4).Value, ws2.Cells(i, 5).Value)
                            .ClearContents
                            .Interior.ColorIndex = 0
                            .Borders.LineStyle = xlLineStyleNone
                        End With
                    '一覧データの場合
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
    
    'メインシートの「セルのロック」を有効にする
    ws1.Cells.Locked = True
    
    Sheets(strDataListSheet).Cells.Clear
        
    ws2.Cells(rowFirstSheetRow, colListInitVal).Value = 0
    
    Set ws1 = Nothing
    Set ws2 = Nothing
        
'正常終了ルーチン
legal_end_rtn:
    '正常終了
    Initialize_Sheet = errNoError
    Exit Function
    
'エラールーチン
error_rtn:
    Initialize_Sheet = process_code
    Exit Function
    
End Function

'背景色初期化
Public Function Initialize_Paint_Sheet(ByVal strItemCheckSheet As String) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn
    
    Dim process_code As Integer
    
    Dim strSheetName As String
    Dim intItemNo As Long
    Dim intListInitVal As Long
                        
    Dim i As Long, j As Long
    Dim lngLastRow As Long
    
    strExecProc = "シート初期化(背景色)"

    '戻り値に不明エラーセット
    process_code = errUnknownError
    Initialize_Paint_Sheet = process_code
            
    '対象シート取得
    strSheetName = GetMainSheetName(strItemCheckSheet)
    '項目数取得
    intItemNo = Sheets(strItemCheckSheet).Cells(rowFirstSheetRow, colItemNo).Value
    
    '一覧部行数取得
    intListInitVal = Sheets(strItemCheckSheet).Cells(rowFirstSheetRow, colListInitVal).Value
        
    If intListInitVal <> 0 Then
        For i = 3 To intItemNo + 2
            DoEvents
            If Sheets(strItemCheckSheet).Cells(i, 2).Value = 1 Then
                Select Case Sheets(strItemCheckSheet).Cells(i, 3).Value
                    '単項データの場合
                    Case 0
                        Sheets(strSheetName).Cells(Sheets(strItemCheckSheet).Cells(i, 4).Value, Sheets(strItemCheckSheet).Cells(i, 5).Value).Interior.ColorIndex = 0
                    '一覧データの場合
                    Case 1
                        For j = Sheets(strItemCheckSheet).Cells(i, 4).Value To intListInitVal + Sheets(strItemCheckSheet).Cells(i, 4).Value - 1
                            Sheets(strSheetName).Cells(j, Sheets(strItemCheckSheet).Cells(i, 5).Value).Interior.ColorIndex = 0
                        Next j
                    Case Else
                End Select
            End If
        Next i
    End If
    
'正常終了ルーチン
legal_end_rtn:
    '正常終了
    Initialize_Paint_Sheet = errNoError
    Exit Function
    
'エラールーチン
error_rtn:
    Initialize_Paint_Sheet = process_code
    Exit Function
    
End Function

'主シート名取得
Public Function GetMainSheetName(ByVal strItemCheckSheet As String) As String
    GetMainSheetName = Sheets(strItemCheckSheet).Cells(rowFirstSheetRow, colSheetName).Value
End Function

'カーソル設定
Public Sub Cursor_Set(ByVal strSheetName As String, ByVal Y As Long, ByVal X As Long)
    Dim strPreChr As String
    
    strPreChr = ""
    If Y > 26 Then
        strPreChr = Chr(cntChrCode_a + (Y \ 26) - 1 + IIf(Y Mod 26 = 0, -1, 0))
        Y = Y Mod 26
    End If
    
    '引数の位置にカーソルを設定
    Application.CutCopyMode = False
    ActiveCell.Copy
    Worksheets(strSheetName).Range(strPreChr & Chr(cntChrCode_a + IIf(Y - 1 > 0, Y - 1, 0)) & X).PasteSpecial xlPasteComments
    Application.CutCopyMode = True
End Sub

'一覧最終行取得
Public Function Items_Last_Row(ByVal strItemCheckSheet As String) As Integer
    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim intAllNo As Long, intAllNoCol As Long, intAllNoRow As Long
    Dim intMaxNo As Long, intNoRow As Long, intFastCol As Long, intLastCol As Long
    Dim Items_Col()
    Dim i As Long

    strExecProc = "一覧最終行取得"
    
    '戻り値に不明エラーセット
    process_code = errUnknownError
    Items_Last_Row = process_code

    '対象シート取得
    strSheetName = GetMainSheetName(strItemCheckSheet)
        
    '総件数位置取得
    intAllNo = Sheets(strItemCheckSheet).Cells.Find("総件数", LookAt:=xlWhole).Row
    intAllNoRow = Sheets(strItemCheckSheet).Cells(intAllNo, 4).Value
    intAllNoCol = Sheets(strItemCheckSheet).Cells(intAllNo, 5).Value

    intMaxNo = 0
    intNoRow = Sheets(strItemCheckSheet).Cells.Find("項番", LookAt:=xlWhole).Row
    intFastCol = Sheets(strItemCheckSheet).Cells(intNoRow, 5).Value
    intLastCol = Sheets(strItemCheckSheet).Cells(Sheets(strItemCheckSheet).Cells(1, 4).Value + 2, 5).Value
    
    '最終行取得
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

    '最終行を項目へ反映
    If intMaxNo > 0 Then
        Sheets(strItemCheckSheet).Cells(1, 8).Value = intMaxNo
        Sheets(strSheetName).Cells(intAllNoRow, intAllNoCol).Value = Format(intMaxNo, "###,###,###,###") & "件"
    Else
        Sheets(strItemCheckSheet).Cells(1, 8).Value = 0
    End If

'正常終了ル ーチン
legal_end_rtn:
    '正常終了
    Items_Last_Row = errNoError
    Exit Function

'エラールーチン
error_rtn:
    Items_Last_Row = process_code
    Exit Function
End Function

'チェックボックス削除
Public Function CheckBox_Delete(ByVal strItemCheckSheet As String) As Integer

    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    Dim strSheetName As String
    Dim obj As Object

    '戻り値に不明エラーセット
    process_code = errUnknownError
    CheckBox_Delete = process_code

    '対象シート取得
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

'正常終了ル ーチン
legal_end_rtn:
    '正常終了
    CheckBox_Delete = errNoError
    Exit Function
'エラールーチン
error_rtn:
    CheckBox_Delete = process_code
    Exit Function
End Function

'操作履歴書込
Public Function OpeLog_Write(ByVal strToolNo As String, ByVal strButtonNo As String) As Integer

    'エラーが発生した場合、エラールーチンへ
    On Error GoTo error_rtn

    '初期化ファイル読み込み
    process_code = Read_InitFile("")

    Dim strUserId As String
    strUserId = Workbooks("マスタメンテナンスメニュー.xlsm").Sheets("マスタメンテナンスメニュー").Cells(2, 12).Value

    'DBオープン
    Dim db As Object
    process_code = DB_Connection(db)
    db.BeginTrans

    'レコードセットオブジェクトの宣言
    Dim rs As Object

    '操作履歴書込
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

    'レコードセットクローズ
    rs.Close

    'DBクローズ
    db.Close

'正常終了ル ーチン
legal_end_rtn:
    OpeLog_Write = errNoError
    Exit Function
'エラールーチン
error_rtn:
    OpeLog_Write = process_code
    Exit Function
End Function

'エラーメッセージ
Public Sub Display_MSG(ByVal errCode As Integer)
    Dim strTTL As String
    Dim strMSG As String

    Select Case errCode
        Case errNoError
            strTTL = "処理終了"
            strMSG = "処理は正常に終了しました"
        Case errPeculiarError
            strTTL = strErrTitle
            strMSG = strErrMsg
        Case errWideError
            strTTL = "全角文字以外の不正入力"
            strMSG = strErrItem & "に全角文字以外の値が入力されています"
        Case errNarrowError
            strTTL = "半角文字以外の不正入力"
            strMSG = strErrItem & "に半角文字以外の値が入力されています"
        Case errNumericError
            strTTL = "数値以外の不正入力"
            strMSG = strErrItem & "が数値ではありません"
        Case errLengthError
            strTTL = "文字数超過"
            strMSG = strErrItem & "の桁数が上限を超えています(上限：" & strErrLength & "桁)"
        Case errLengthLessError
            strTTL = "文字数過少"
            strMSG = strErrItem & "の桁数が規定に満たない値です(規定：" & strErrLength & "桁)"
        Case errDBConnectionhError
            strTTL = "データベース接続エラー"
            strMSG = "データベースへの接続に失敗しました"
        Case errDBUpdateError
            strTTL = "データベース更新エラー"
            strMSG = strErrMsg
        Case errDBInsertError
            strTTL = "データベース追加エラー"
            strMSG = strErrMsg
        Case errDBDeleteError
            strTTL = "データベース削除エラー"
            strMSG = strErrMsg
        Case errRecordSetError
            strTTL = "データ抽出エラー"
            strMSG = "データ抽出に失敗しました"
        Case errRecordSetSetError
            strTTL = "データ作成エラー"
            strMSG = "データは抽出しましたが、シートへの展開に失敗しました"
        Case errNoDataError
            strTTL = "データ無し"
            strMSG = "該当するデータは存在しません"
        Case errIllegalInputError
            strTTL = "不正入力エラー"
            strMSG = strExecProc & "の処理中に、不正な入力が存在するため、処理を中断しました"
        Case errNoInputError
            strTTL = "未入力エラー"
            strMSG = strExecProc & "の処理中に、未入力項目が検出されたため、" & Chr(&HD) & "処理を中断しました"
        Case errNoOutputError
            strTTL = "出力エラー"
            strMSG = "出力するデータが存在しません"
        Case errIniFileReadError
            strTTL = "INIファイル読み込みエラー"
            strMSG = "INIファイルの読み込みに失敗しました"
        Case errUnknownError
            strTTL = "異常終了"
            strMSG = strExecProc & "の処理中に異常終了しました"
        Case errCSVFileOpenError
            strTTL = "CSV読み込みエラー"
            strMSG = "CSVファイルの読み込みに失敗しました"
        Case errCSVFileOutputError
            strTTL = "CSV出力エラー"
            strMSG = "CSVファイルの出力に失敗しました"
        Case errCSV_FileDeleteError
            strTTL = "CSV削除エラー"
            strMSG = "CSVファイルの削除に失敗しました"
        Case errText_FileOpenError
            strTTL = "テキストファイル読込エラー"
            strMSG = "テキストファイルの読込に失敗しました"
        Case errText_FileOutputError
            strTTL = "テキストファイル出力エラー"
            strMSG = "テキストファイルの出力に失敗しました"
        Case errNoInputDataError
            strTTL = "読込エラー"
            strMSG = "ファイルにデータが存在しません"
    End Select

    MsgBox strMSG, vbOKOnly, strTTL

End Sub
