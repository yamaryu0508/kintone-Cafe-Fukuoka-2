'' インターネット接続状況を判定するライブラリ等の準備
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
  Private Const INTERNET_CONNECTION_MODEM As Long = &H1
  Private Const INTERNET_CONNECTION_LAN As Long = &H2
  Private Const INTERNET_CONNECTION_PROXY As Long = &H4
  Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20

'' Book全体の定数の宣言
' kintoneアクセス関係
Public Const DOMAIN_NAME As String = "[subdomain].cybozu.com" ' ドメイン名「[subdomain].cybozu.com」
Public Const KINTONE_API_BASE_URI As String = "https://" & DOMAIN_NAME & "/k/v1/" ' kintone REST APIリクエスト用ベースURI
Public Const PORT_NUM As String = "443" ' HTTPSポート番号
Public Const CUSTOMER_INFO_APP_ID As Integer = [app id] ' 顧客情報アプリのアプリID
Public Const CUSTOMER_INFO_API_TOKEN As String = "[API Token]" ' 顧客情報アプリのAPIトークン
Public Const ITEMS_MANAGE_APP_ID As Integer = [app id] ' 物品情報アプリのアプリID
Public Const ITEMS_MANAGE_API_TOKEN As String = "[API Token]" ' 物品情報アプリのAPIトークン
' 表関係
Public Const TABLE_ROW_START As Integer = 27 ' 表の開始行
Public Const TABLE_COL_START As Integer = 3 ' 表の開始列
Public Const TABLE_SIZE As Integer = 10 ' 表のサイズ（最大行数）


'' 物品貸出手続き
Sub ApplyRental()
  If IsOnline() = False Then ' オフライン時はメッセージを表示して終了
    ShowErrMsg ("OFFLINE")
    Exit Sub
  End If
    changeStatus ("貸出中") ' 物品情報アプリのステータスを貸出用に変更
End Sub

'' 物品返却手続き
Sub StopRental()
  If IsOnline() = False Then ' オフライン時はメッセージを表示して終了
    ShowErrMsg ("OFFLINE")
    Exit Sub
  End If
    changeStatus ("未使用") ' 物品情報アプリのステータスを返却用に変更
End Sub

Function changeStatus(ByVal strStatus As String)
    
  '' 変数群の宣言
  Dim objHttpRequest As Object ' XMLHTTPオブジェクト（関数返り値）
  Dim objHeaders As Object ' HTTPリクエストヘッダ
  Dim strUri As String ' リクエストURI
  Dim strJSON As String ' リクエストボディ用JSON
  Dim strOwner As String ' 提供先・設置場所
  Dim count As Integer ' カウンタ
  Dim rowCount As Integer ' 列カウンタ
  
  Dim sheetRental As Worksheet ' ワークシート
  Set sheetRental = ThisWorkbook.Worksheets("トライアルサービス申込書")
        
  '' 提供先・設置場所として、シート中のお客さま名をセット
  If (strStatus = "貸出中") Then
     strOwner = sheetRental.Range("L12").value
  ElseIf (strStatus = "未使用") Then
     strOwner = ""
  End If
  
  '' JSON文字列を作成
  strJSON = "{"
  strJSON = strJSON & """app"": " & ITEMS_MANAGE_APP_ID & ","
  strJSON = strJSON & """records"": ["
    
  For count = 0 To (TABLE_SIZE - 1) ' レコード部分を作成
    rowCount = TABLE_ROW_START + count
    If (sheetRental.Range("C" & rowCount).value <> "") Then
      strJSON = strJSON & "{"
      strJSON = strJSON & """id"": " & sheetRental.Range("BE" & rowCount).value & ","
      strJSON = strJSON & """record"": {"
      strJSON = strJSON & """operationStatus"": {""value"": """ & strStatus & """},"
      strJSON = strJSON & """ownerLocation"": {""value"": """ & strOwner & """}"
      strJSON = strJSON & "}"
      strJSON = strJSON & "},"
    End If
  Next count
    
  strJSON = Left(strJSON, Len(strJSON) - 1) ' 最後のレコードのコンマ「,」を削除
    
  strJSON = strJSON & "]"
  strJSON = strJSON & "}"
  
  Debug.Print strJSON
    
  '' HTTPリクエスト
  strUri = KINTONE_API_BASE_URI & "records.json" ' リクエストURIを作成
  Set objHeaders = CreateObject("Scripting.Dictionary") ' リクエストヘッダを作成
  objHeaders.Add "X-Cybozu-API-Token", ITEMS_MANAGE_API_TOKEN ' リクエストヘッダにAPIトークンを追加
  objHeaders.Add "Host", DOMAIN_NAME + ":" + PORT_NUM ' リクエストヘッダにホスト名を追加
  objHeaders.Add "Content-Type", "application/json" ' リクエストヘッダに「Content-Type:application/json」を追加
  Set objHttpRequest = requestHttp("PUT", strUri, objHeaders, strJSON) ' リクエスト送信
  
  If objHttpRequest.status = 200 Then ' レスポンスに対する処理
    If (strStatus = "貸出中") Then
      MsgBox "貸出手続き完了"
    ElseIf (strStatus = "未使用") Then
      MsgBox "返却手続き完了"
    End If
  Else
      MsgBox objHttpRequest.responseText
  End If
    
  Set objHttpRequest = Nothing ' オブジェクト解放
  Set objHeaders = Nothing ' オブジェクト解放
  Set sheetRental = Nothing ' オブジェクト解放
  
End Function

'' JSONパース用関数
Public Function parseJSON(strJSON As String) As Object
  Dim lib As New JSONLib 'Instantiate JSON class object
  Set parseJSON = lib.parse(CStr(strJSON))
End Function

'' URIエンコード用関数
Public Function URI_Encode(ByVal strOrg As String) As String
  Dim d As Object
  Dim elm As Object
   
  strOrg = Replace(strOrg, "\", "\\")
  strOrg = Replace(strOrg, "'", "\'")
  Set d = CreateObject("htmlfile")
  Set elm = d.createElement("span")
  elm.setAttribute "id", "result"
  d.appendChild elm
  d.parentWindow.execScript "document.getElementById('result').innerText = encodeURIComponent('" & strOrg & "');", "JScript"
  URI_Encode = elm.innerText
End Function

'' HTTPリクエスト用関数
Public Function requestHttp(strMethod As String, strUri As String, ByVal dictHeaders As Object, varReqBody As Variant) As Object
  Dim header As Variant ' ヘッダ（カウント用）
  Dim objRequestHttp As Object ' オブジェクトの作成
  Set objRequestHttp = CreateObject("MSXML2.XMLHTTP")
  objRequestHttp.Open strMethod, strUri, False
  For Each header In dictHeaders.keys ' ヘッダの追加
    objRequestHttp.setRequestHeader header, dictHeaders.Item(header)
  Next
  objRequestHttp.setRequestHeader "If-Modified-Since", "Thu, 01 Jun 1970 00:00:00 GMT"
  objRequestHttp.send (varReqBody) ' ボディのセット
  Set requestHttp = objRequestHttp
End Function

'' インターネット接続状況を確認する関数
Function IsOnline() As Boolean
  Dim L As Long
  Dim R As Long
  R = InternetGetConnectedState(L, 0&)
  If R = 0 Then
    IsOnline = False
  Else
    If R <= 4 Then
      IsOnline = True
    Else
      IsOnline = False
    End If
  End If
End Function

'' エラーメッセージを表示する関数
Function ShowErrMsg(strType As String)
  Select Case strType
    Case "OFFLINE" ' オフライン時のメッセージ
      MsgBox "PC is offline" & vbCrLf & "(Cannot access kintone)", vbExclamation
    Case "NO_RECORD" ' ルックアップ失敗（レコードなし）時のメッセージ
      MsgBox "Lookup from kintone failed" & vbCrLf & "(No appropriate record exists)", vbExclamation
    End Select
End Function

