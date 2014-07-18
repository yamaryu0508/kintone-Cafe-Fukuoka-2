Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range) ' シート変更プロージャ
  
  '' オフライン時のエラー処理
  If IsOnline() = False Then ' オフライン時はメッセージを表示して終了
    ShowErrMsg ("OFFLINE")
    Exit Sub
  End If
  
  '' 変数群の宣言
  Dim objHttpRequest As Object ' XMLHTTPオブジェクト（関数返り値）
  Dim objHeaders As Object ' HTTPリクエストヘッダ
  Dim strUri As String ' リクエストURI
  Dim strQuery As String ' リクエストURI中のクエリ
  Dim strJSON As String ' リクエストボディ用JSON
  Dim objJSON As Object ' レスポンスボディ用JSON
  Dim record As Variant ' レコード（カウント用）
  Dim rangeTarget As Range ' 管理番号範囲（カウント用）
  Dim rangeTargets As Range ' 管理番号範囲
  
  Dim sheetRental As Worksheet ' ワークシート
  Set sheetRental = ThisWorkbook.Worksheets("トライアルサービス申込書")
  
  '' 管理番号から機器情報をルックアップする処理のブロック
  If (Target.Column = TABLE_COL_START) And (Target.Row >= TABLE_ROW_START And Target.Row <= (TABLE_ROW_START + TABLE_SIZE)) Then
    '' 管理番号を抽出し、ルックアップする
    Set rangeTargets = sheetRental.Range(Target.Address)
    For Each rangeTarget In rangeTargets ' 管理番号が範囲で渡されることを想定
      If rangeTarget.Column = TABLE_COL_START Then ' 結合セルに対してすべてEachしてしまうため、3列目に限定
        If rangeTarget.value <> "" Then ' 管理番号入力欄に値が入力されたときの処理
          '' HTTPリクエスト
          strQuery = "id = " & """" & rangeTarget.value & """ limit 1" 'クエリ設定
          strQuery = URI_Encode(strQuery) ' クエリのURIエンコード
          strUri = KINTONE_API_BASE_URI & "records.json?app=" & ITEMS_MANAGE_APP_ID & "&query=" & strQuery ' リクエストURI
          
          Set objHeaders = CreateObject("Scripting.Dictionary") ' リクエストヘッダを作成
          objHeaders.Add "X-Cybozu-API-Token", ITEMS_MANAGE_API_TOKEN ' リクエストヘッダにAPIトークンを追加
          objHeaders.Add "Host", DOMAIN_NAME + ":" + PORT_NUM ' リクエストヘッダにホスト名を追加
          Set objHttpRequest = requestHttp("GET", strUri, objHeaders, Null) ' リクエスト送信

          If objHttpRequest.status = 200 Then ' レスポンスに対する処理
            strJSON = objHttpRequest.responseText ' レスポンスボディをセット
            Debug.Print strJSON
            Set objJSON = parseJSON(strJSON) ' レスポンスボディに対してJSONパース
            If objJSON("records").count = 0 Then ' レスポンスが空の場合はメッセージを表示して終了
              ShowErrMsg ("NO_RECORD")
              Exit Sub
            End If
            ActiveSheet.Unprotect ' セルに値を設定するため、一時的にシート保護解除
            For Each record In objJSON("records") ' セルに値をセット
              sheetRental.Range("I" & rangeTarget.Row).value = record.Item("type").Item("value")
              sheetRental.Range("Q" & rangeTarget.Row).value = record.Item("maker").Item("value")
              sheetRental.Range("Y" & rangeTarget.Row).value = record.Item("model").Item("value")
              sheetRental.Range("BE" & rangeTarget.Row).value = record.Item("recordNum").Item("value")
            Next record
            ActiveSheet.Protect ' シート保護復帰
          Else
            MsgBox strJSON, vbCritical
          End If
    
          Set objHttpRequest = Nothing ' オブジェクト解放
          Set objHeaders = Nothing ' オブジェクト解放
          Set objJSON = Nothing ' オブジェクト解放
        Else ' 管理番号入力欄が空白になった時の処理
          ActiveSheet.Unprotect ' セルに値を設定するため、一時的にシート保護解除
          sheetRental.Range("I" & rangeTarget.Row).value = ""
          sheetRental.Range("Q" & rangeTarget.Row).value = ""
          sheetRental.Range("Y" & rangeTarget.Row).value = ""
          sheetRental.Range("BE" & rangeTarget.Row).value = ""
          sheetRental.Range("AI" & rangeTarget.Row).value = ""
          ActiveSheet.Protect ' シート保護復帰
        End If ' 管理番号入力欄に対する分岐
      End If ' 結合セル用の対策部分
    Next ' rangeTarget
  End If ' 表部分のルックアップ

  
  '' 会社コードから会社名等をルックアップする処理のブロック
  If (Target.Column = 57 And Target.Row = 12) Then
    If Target.value <> "" Then ' 会社コード入力欄に値が入力されたとき
      Set rangeTargets = sheetRental.Range(Target.Address)
      For Each rangeTarget In rangeTargets ' 複数セルでコピーされたときを想定し、対象Rangeを順番に処理
        '' HTTPリクエスト
        strQuery = "companyCode = " & """" & rangeTarget.value & """limit 1" ' クエリの設定
        strQuery = URI_Encode(strQuery) ' クエリのURIエンコード
        strUri = KINTONE_API_BASE_URI & "records.json?app=" & CUSTOMER_INFO_APP_ID & "&query=" & strQuery ' URIエンコード
        
        Set objHeaders = CreateObject("Scripting.Dictionary") ' リクエストヘッダを作成
        objHeaders.Add "X-Cybozu-API-Token", CUSTOMER_INFO_API_TOKEN ' リクエストヘッダにAPIトークンを追加
        objHeaders.Add "Host", DOMAIN_NAME + ":" + PORT_NUM ' リクエストヘッダにホスト名を追加
        Set objHttpRequest = requestHttp("GET", strUri, objHeaders, Null) ' リクエスト送信
    
        If objHttpRequest.status = 200 Then ' レスポンスに対する処理
          strJSON = objHttpRequest.responseText ' レスポンスボディをセット
          Debug.Print strJSON
          Set objJSON = parseJSON(strJSON) ' レスポンスボディに対してJSONパース
          If objJSON("records").count = 0 Then ' レスポンスが空の場合はメッセージを表示して終了
            ShowErrMsg ("NO_RECORD")
            Exit Sub
          End If
          ActiveSheet.Unprotect ' セルに値を設定するため、一時的にシート保護解除
          For Each record In objJSON("records")
            sheetRental.Range("L12").value = record.Item("companyName").Item("value")
            sheetRental.Range("L15").value = "〒" & record.Item("zipCode").Item("value") & " " & record.Item("Address").Item("value")
            sheetRental.Range("S18").value = record.Item("inCharge").Item("value")
            sheetRental.Range("AC18").value = record.Item("tel").Item("value")
            sheetRental.Range("AO18").value = record.Item("mail").Item("value")
          Next record
          ActiveSheet.Protect ' シート保護復帰
        Else
          MsgBox strJSON, vbCritical
        End If
    
        Set objHttpRequest = Nothing ' オブジェクト解放
        Set objHeaders = Nothing ' オブジェクト解放
        Set objJSON = Nothing ' オブジェクト解放
      Next ' rangeTarget
    Else ' 会社コード入力欄が空白になったとき
      ActiveSheet.Unprotect ' セルに値を設定するため、一時的にシート保護解除
      sheetRental.Range("L12").value = ""
      sheetRental.Range("L15").value = ""
      sheetRental.Range("S18").value = ""
      sheetRental.Range("AC18").value = ""
      sheetRental.Range("AO18").value = ""
      ActiveSheet.Protect 'シート保護復帰
    End If ' 会社コード入力欄に対する分岐
  End If ' 会社コード部分のルックアップ
End Sub
