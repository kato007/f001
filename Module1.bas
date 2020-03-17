Attribute VB_Name = "Module1"
Const SHEET = "取り込み結果"
Const MAXCELL = 46
Public Type haifuData
    N As String
    N_name As String
End Type
Dim TUDUKI_LINE As Integer
'** 29  扱い先
Public Sub makeSeikyusyoOfNikkaHome()
    
    'アプリケーション形式を変換
    With Application
        .ReferenceStyle = xlR1C1
    End With
    '***シートデータをクリア
    Call clearData
    '***データを読み込む
    Call makeNikkaHomeSeikyuData
    '***0648M020の枝番をまとめる
    Call chgdata("データ")
    '***データを読み込む
    Call makeNikkaHomeSeikyuData_0648MA1X
    '*** データの並べ替え
    Call sortData
    '***シートを削除する（前回配布データを削除）
    Call initDelSheets
    '***配布先を取得
    'Call gethaifusaki(haifusaki)
    '***配布用シートを得意先毎に作成
    Call makeSheet
    With Application
        .ReferenceStyle = xlA1
    End With
    '***各シートの不要部分を削除
    Call DelSpaceArea
    Call OutputSheetFile
    MsgBox "終わりました"
End Sub


'***配布用シートを工場別に作成
Private Sub makeSheet()
    Dim i As Integer
    Dim tokuiCd As String
    Dim intLine As Integer  '***データのカウンター
    Dim intSline As Integer '***表紙の行カウンター
    Dim intUline As Integer '***内訳の行カウンター
    Dim subTotal As Double    '***明細の合計格納
    Dim strKeyKojino As String  '***比較キーの工事ＮＯ
    
    intLine = 2
    
    With Sheets("データ")
        
        Do While .Cells(intLine, 1) <> ""
            If Trim(tokuiCd) <> Trim(.Cells(intLine, 5)) Then
                tokuiCd = .Cells(intLine, 5)
                Sheets("表紙").Select
                Sheets("表紙").Copy After:=Sheets(4)
                ActiveSheet.Name = tokuiCd & "表紙"
                Sheets("内訳").Select
                Sheets("内訳").Copy After:=Sheets(4)
                ActiveSheet.Name = tokuiCd & "内訳"
                Call initset(tokuiCd & "表紙", intLine)
                intSline = 15
                intUline = 6
                tokuiCd = .Cells(intLine, 5)
            End If
            '**表紙
            Call setData(tokuiCd & "表紙", intLine, intSline)
            '***明細
            Call setUchiwakeData(tokuiCd & "内訳", intLine, intUline, strKeyKojino, subTotal)
            intLine = intLine + 1
        Loop

    End With

End Sub

'****************************************************************************
'***初期値をセットする（表紙）
'****************************************************************************
Private Sub initset(ByVal sheetName As String, ByVal intLine As Integer)
    Dim intLineTitle As Integer
    
    Sheets(sheetName).Cells(8, 2) = Sheets("データ").Cells(intLine, 6)
    Sheets(sheetName).Cells(5, 18) = Format(DateAdd("M", -1, Date), "MM")
    
    '***タイトル
    If InStr(Sheets("データ").Cells(intLine, 6), "大問屋") > 0 Then
       Sheets(sheetName).Cells(6, 1) = "大問屋　株式会社　御中"
       Sheets(sheetName).Cells(8, 2) = Replace(Sheets("データ").Cells(intLine, 6), "大問屋㈱", "")
    Else
        If Mid(Sheets("データ").Cells(intLine, 6), 1, 8) = "ニッカホーム関東" Then
            Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社　御中"
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets("データ").Cells(intLine, 6) & "営業所", "㈱", "")
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets(sheetName).Cells(8, 2), "関東", "")
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets(sheetName).Cells(8, 2), "ニッカホーム", "")
        Else
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets("データ").Cells(intLine, 6), "㈱", "")
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets(sheetName).Cells(8, 2), "ニッカホーム", "")
        End If
    End If
    intLineTitle = 30
    Debug.Print Sheets("データ").Cells(intLine, 5)
    If Sheets("データ").Cells(intLine, 5) = "0348M570" Then
        Debug.Print
    End If
    

    Do While Sheets("メニュー").Cells(intLineTitle, 2) <> ""
        If Sheets("メニュー").Cells(intLineTitle, 2) = Sheets("データ").Cells(intLine, 5) Then
            If Sheets("メニュー").Cells(intLineTitle, 1) <> "" Then
                Sheets(sheetName).Cells(6, 1) = Sheets("メニュー").Cells(intLineTitle, 1) & " 御中"
            End If
            Exit Sub
        End If
        intLineTitle = intLineTitle + 1
    Loop


    If Sheets("データ").Cells(intLine, 5) = "1148M870" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム福岡㈱　御中"
       Sheets(sheetName).Cells(8, 2) = "福岡南営業所"
    End If

    If Sheets("データ").Cells(intLine, 5) = "1148M880" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム福岡㈱　御中"
       Sheets(sheetName).Cells(8, 2) = "福岡筑紫（営"
    End If
    If Sheets("データ").Cells(intLine, 5) = "1148M890" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム福岡㈱　御中"
       Sheets(sheetName).Cells(8, 2) = "福岡東営業所"
    End If
    If Sheets("データ").Cells(intLine, 5) = "1148M900" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム福岡㈱　御中"
       Sheets(sheetName).Cells(8, 2) = "福岡早良（営"
    End If

    If Sheets("データ").Cells(intLine, 5) = "0348ME30" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "大問屋　南横浜店"
    End If
    
    If Sheets("データ").Cells(intLine, 5) = "0348ME50" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "大問屋　西東京店"
    End If

    '***20200221
    If Sheets("データ").Cells(intLine, 5) = "0348ME60" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "大型事業部"
    End If
    
    If Sheets("データ").Cells(intLine, 5) = "0348ME50" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "大問屋　西東京店"
    End If

    If Sheets("データ").Cells(intLine, 5) = "0348M710" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "横浜保土ヶ谷営業所"
    End If


    If Sheets("データ").Cells(intLine, 5) = "0348M690" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "本社"
    End If

End Sub

'****************************************************************************
'***データをセットする（表紙）
'****************************************************************************
Private Sub setData(ByVal sheetName As String, ByVal intLine As Integer, ByRef intDummy As Integer)
    Dim aTenmei As Variant
    Dim i As Integer
    Dim intSline As Integer
    Dim btrue As Boolean
    Dim atenmeiCHK As String

    intSline = 16
    btrue = False

    '***最初の1行の金額欄が空白じゃなかったら、同じ工事番号を探す＋現場名も一致すればＯＫ
    '*cells(intSline,11) は金額欄
    '*cells(intline,20) は注番
    Do While Sheets(sheetName).Cells(intSline, 11) <> ""
        If Trim(Sheets(sheetName).Cells(intSline, 1)) = Trim(Sheets("データ").Cells(intLine, 20)) Then
            Sheets("データ").Cells(intLine, 29) = Replace(Sheets("データ").Cells(intLine, 29), "*", "/")    '**9/1更新
            If InStr(Sheets("データ").Cells(intLine, 54) & Sheets("データ").Cells(intLine, 55), "／") > 0 Then
                    aTenmei = Split(Sheets("データ").Cells(intLine, 54) & Sheets("データ").Cells(intLine, 55), "／")
                    atenmeiCHK = aTenmei(1)
            Else
                If InStr(Sheets("データ").Cells(intLine, 29), "/") > 0 Then
                    aTenmei = Split(Sheets("データ").Cells(intLine, 29), "/")
                    atenmeiCHK = aTenmei(1)
                Else
                    atenmeiCHK = ""                                                                                 '**現場名
                End If
            End If
            '*現場名が同じか？同じなら金額を加算する。
            '*現場名から工事番号に変更する
            'If Trim(Sheets(sheetName).Cells(intSline, 3)) = Trim(atenmeiCHK) Then
            If Trim(Sheets(sheetName).Cells(intSline, 1)) = Trim(Sheets("データ").Cells(intLine, 20)) Then
                Sheets(sheetName).Cells(intSline, 11) = Sheets(sheetName).Cells(intSline, 11) + Sheets("データ").Cells(intLine, 19)
                btrue = True
            End If
        End If
        intSline = intSline + 1
        If intSline = 37 Then
            intSline = 42
        End If
        If intSline = 72 Then
            intSline = 77
        End If
        If intSline = 72 Then
            intSline = 77
        End If
        If intSline = 108 Then
            intSline = 111
        End If
    Loop
    
    If Not (btrue) Then
        Sheets(sheetName).Cells(intSline, 1) = Trim(Sheets("データ").Cells(intLine, 20))                        '**工事番号（注番）
        Sheets("データ").Cells(intLine, 29) = Replace(Sheets("データ").Cells(intLine, 29), "*", "/")    '**9/1更新
        If InStr(Sheets("データ").Cells(intLine, 54) & Sheets("データ").Cells(intLine, 55), "／") > 0 Then
                aTenmei = Split(Sheets("データ").Cells(intLine, 54) & Sheets("データ").Cells(intLine, 55), "／")
                Sheets(sheetName).Cells(intSline, 3) = aTenmei(1)                                                   '**現場名
                Sheets(sheetName).Cells(intSline, 8) = aTenmei(0)                                                   '**担当
        Else
            If InStr(Sheets("データ").Cells(intLine, 29), "/") > 0 Then
                aTenmei = Split(Sheets("データ").Cells(intLine, 29), "/")
                Sheets(sheetName).Cells(intSline, 3) = aTenmei(1)                                                   '**現場名
                Sheets(sheetName).Cells(intSline, 8) = aTenmei(0)                                                   '**担当
            Else
                Sheets(sheetName).Cells(intSline, 8) = Sheets("データ").Cells(intLine, 29)                          '**現場名
            End If
        End If
        Sheets(sheetName).Cells(intSline, 11) = Sheets("データ").Cells(intLine, 19)                             '**金額
    End If

    If Sheets("データ").Cells(intLine, 5) = "1148M870" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム福岡㈱　御中"
       Sheets(sheetName).Cells(8, 2) = "福岡南営業所"
    End If

    If Sheets("データ").Cells(intLine, 5) = "1148M880" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム福岡㈱　御中"
       Sheets(sheetName).Cells(8, 2) = "福岡筑紫（営"
    End If
    If Sheets("データ").Cells(intLine, 5) = "1148M890" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム福岡㈱　御中"
       Sheets(sheetName).Cells(8, 2) = "福岡東営業所"
    End If
    If Sheets("データ").Cells(intLine, 5) = "1148M900" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム福岡㈱　御中"
       Sheets(sheetName).Cells(8, 2) = "福岡早良（営"
    End If
    
    If Sheets("データ").Cells(intLine, 5) = "0348ME30" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "大問屋 南横浜店"
    End If
    
    If Sheets("データ").Cells(intLine, 5) = "0348ME50" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "大問屋 西東京店　"
    End If
    '***20200221
    If Sheets("データ").Cells(intLine, 5) = "0348ME60" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "大型事業部"
    End If
    
    If Sheets("データ").Cells(intLine, 5) = "0348ME50" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "大問屋　西東京店"
    End If

    If Sheets("データ").Cells(intLine, 5) = "0348M710" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "横浜保土ヶ谷営業所"
    End If


    If Sheets("データ").Cells(intLine, 5) = "0348M690" Then
       Sheets(sheetName).Cells(6, 1) = "ニッカホーム関東株式会社 御中"
       Sheets(sheetName).Cells(8, 2) = "本社"
    End If
End Sub

'****************************************************************************
'***データをセットする（内訳・明細）
'****************************************************************************
Private Sub setUchiwakeData(ByVal sheetName As String, _
                            ByVal intLine As Integer, _
                            ByRef intUline As Integer, _
                            ByRef strKeyKojino As String, _
                            ByRef subTotal As Double)
    
    Dim aTenmei As Variant
    '***工事№
    
    subTotal = subTotal + Sheets("データ").Cells(intLine, 18) * Sheets("データ").Cells(intLine, 17)
    Sheets(sheetName).Cells(intUline, 2) = Trim(Sheets("データ").Cells(intLine, 20))                              '**工事番号（注番
    Sheets(sheetName).Cells(intUline, 1) = Mid(Sheets("データ").Cells(intLine, 21), 3, 2) & "/" _
                                         & Mid(Sheets("データ").Cells(intLine, 21), 5, 2)                       '***出荷日
    Sheets(sheetName).Cells(intUline + 1, 5) = Sheets("データ").Cells(intLine, 17)                              '***数量
    Sheets(sheetName).Cells(intUline + 1, 6) = Sheets("データ").Cells(intLine, 18)                              '***単価
    '***集計をいれる(工事番号,日付)
    If strKeyKojino <> Trim(Sheets("データ").Cells(intLine + 1, 20)) Or _
        Trim(Sheets("データ").Cells(intLine + 1, 5)) <> Trim(Sheets("データ").Cells(intLine, 5)) Or _
        Trim(Sheets("データ").Cells(intLine + 1, 21)) <> Trim(Sheets("データ").Cells(intLine, 21)) Then
        If subTotal <> 0 Then
            Sheets(sheetName).Cells(intUline + 1, 8) = subTotal
        End If
        subTotal = 0
        strKeyKojino = Trim(Sheets("データ").Cells(intLine + 1, 20))
    End If
    '***台・個を判定（図番があるなし）
    If Trim(Sheets("データ").Cells(intLine, 47)) = "" Then
        Sheets(sheetName).Cells(intUline, 5) = "台"
        Sheets(sheetName).Cells(intUline + 1, 2) = Sheets("データ").Cells(intLine, 15)
    Else
        Sheets(sheetName).Cells(intUline, 5) = "個"
        Sheets(sheetName).Cells(intUline + 1, 2) = Sheets("データ").Cells(intLine, 15)
    End If
    
    '***現場名と担当名(2019.9.2 備考の内容を印字するように変更)
    If InStr(Sheets("データ").Cells(intLine, 54) & Sheets("データ").Cells(intLine, 55), "／") > 0 Then
        aTenmei = Split(Sheets("データ").Cells(intLine, 54) & Sheets("データ").Cells(intLine, 55), "／")
        Sheets(sheetName).Cells(intUline, 3) = aTenmei(1)                                                       '**現場名
        Sheets(sheetName).Cells(intUline, 4) = aTenmei(0)                                                       '**担当
    Else
        If InStr(Sheets("データ").Cells(intLine, 29), "/") > 0 Then
            aTenmei = Split(Sheets("データ").Cells(intLine, 29), "/")
            Sheets(sheetName).Cells(intUline, 3) = aTenmei(1)                                                       '**現場名
            Sheets(sheetName).Cells(intUline, 4) = aTenmei(0)                                                       '**担当
        Else
            Sheets(sheetName).Cells(intUline, 4) = Sheets("データ").Cells(intLine, 29)                              '**現場名
        End If
    End If
    intUline = intUline + 2
    If intUline = 56 Then
        intUline = 61
    End If

    If intUline = 111 Then
        intUline = 116
    End If

    If intUline = 166 Then
        intUline = 171
    End If
    
    If intUline = 221 Then
        intUline = 226
    End If
    If intUline = 276 Then
        intUline = 281
    End If
    If intUline = 331 Then
        intUline = 336
    End If
    If intUline = 386 Then
        intUline = intUline + 5
    End If
    If intUline = 441 Then
        intUline = intUline + 5
    End If
    If intUline = 496 Then
        intUline = intUline + 5
    End If
    If intUline = 551 Then
        intUline = intUline + 5
    End If
    'If intUline > 386 Then
        'If intUline Mod 55 = 0 Then
            'intUline = intUline + 5
        'End If
    'End If
        


End Sub
'****************************************************************************
'***配布先シートに貼り付ける
'****************************************************************************
Private Sub PasteFactToSheet(ByVal sheetName As String, ByVal stN As String)
    Dim iRow As Integer
    Dim iColumn As Integer
    
    With Sheets(SHEET)
        .Cells(1, 1).AutoFilter Field:=3, Criteria1:=stN
        iColumn = .Cells(1, 1).End(xlToRight).Column
        iRow = .Cells(1, 1).End(xlDown).Row
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).Copy
        Sheets(sheetName).Paste
        '***配布先シートに貼り付けたあとに、関数を入れていく
        Call insFunc(sheetName)
    End With
End Sub
'****************************************************************************
'***配布先シートに貼り付けたあとに、関数を入れていく
'****************************************************************************
Private Sub insFunc(ByVal sheetName As String)
    Dim iniLine As Integer
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim i As Integer
    
    intLine = 1
    
    With Sheets(sheetName)
        iColumn = .Cells(1, 1).End(xlToRight).Column
        iRow = .Cells(1, 1).End(xlDown).Row
        '***一旦色をリセットする
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).Interior.ColorIndex = xlNone
    
        Do While .Cells(intLine, 1) <> ""
            If .Cells(intLine, 13) = "先行確認" Then
                '***色をつける
                .Range(.Cells(intLine, 1), .Cells(intLine, MAXCELL)).Interior.ColorIndex = 35
                For i = 0 To 30
                    .Cells(intLine, 15 + i) = "=RC[-1]+R[-3]C-R[-2]C-R[-1]C"
                Next i

            End If
            intLine = intLine + 1
        Loop
        '***条件付書式設定をする
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).FormatConditions.Delete
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="0"
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).FormatConditions(1).Interior.ColorIndex = 7
        .Cells.EntireColumn.AutoFit
    End With

End Sub
'****************************************************************************
'***工場配布用のシートを削除する
'****************************************************************************
Private Sub initDelSheets()
    Dim ws As Worksheet
    '***削除計画をOFFにする
    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name <> "表紙" And ws.Name <> "内訳" And ws.Name <> "データ" And ws.Name <> "メニュー" Then
            Sheets(ws.Name).Delete
        End If
    Next ws
    '***削除計画をONにする
    Application.DisplayAlerts = True

End Sub
'***配布先を取得
Private Sub gethaifusaki(ByRef haifusaki() As haifuData)
    
    Dim i As Integer
    
    i = 0
    With Sheet2
        Do While .Cells(17 + i, 7) <> ""
           ReDim Preserve haifusaki(i)
            haifusaki(i).N = .Cells(17 + i, 7)
            haifusaki(i).N_name = .Cells(17 + i, 7 + 1)
            i = i + 1
        Loop
    End With
End Sub
'*************************************************************************************************
'シート毎にデータ件数に合わせてページを設定する（いらないところを消す）
'*************************************************************************************************
Public Sub DelSpaceArea()
    Dim intIdx As Integer       '処理用インデックス
    Dim intWksCnt As Integer    '処理用カウンタ
    
    'シート数取得
    intWksCnt = Excel.ActiveWorkbook.Worksheets.Count
    '***OFFにする
    Application.DisplayAlerts = False
    
    'シート数ループ
    For intIdx = 1 To intWksCnt
        '対象シート名取得
        strWksnme = Worksheets(intIdx).Name
        If Right(strWksnme, 2) = "表紙" And strWksnme <> "表紙" Then
            If Sheets(strWksnme).Cells(42, 1) = "" Then
                Call Setpage1(strWksnme)
            Else
                If Sheets(strWksnme).Cells(77, 1) = "" Then
                    Call Setpage2(strWksnme)
                Else
                    If Sheets(strWksnme).Cells(111, 1) = "" Then
                        Call Setpage3(strWksnme)
                    Else
                        If Sheets(strWksnme).Cells(145, 1) = "" Then
                            Call Setpage4(strWksnme)
                        End If
                    End If

                End If
            End If

        End If
        If Right(strWksnme, 2) = "内訳" And strWksnme <> "内訳" Then
            '１ページ
            If Sheets(strWksnme).Cells(61, 1) = "" Then
                Call SetpageUchiwake1(strWksnme)
            Else
                '２ページ
                If Sheets(strWksnme).Cells(116, 1) = "" Then
                    Call SetpageUchiwake2(strWksnme)
                Else
                    '３ページ
                    If Sheets(strWksnme).Cells(171, 1) = "" Then
                        Call SetpageUchiwake3(strWksnme)
                     Else
                        '4ページ
                        If Sheets(strWksnme).Cells(226, 1) = "" Then
                            Call SetpageUchiwake4(strWksnme)
                        Else
                        '5ページ
                            If Sheets(strWksnme).Cells(281, 1) = "" Then
                                Call SetpageUchiwake5(strWksnme)
                            Else
                            '6ページ
                            If Sheets(strWksnme).Cells(336, 1) = "" Then
                                Call SetpageUchiwake6(strWksnme)
                            Else
                            '7ページ
                            If Sheets(strWksnme).Cells(391, 1) = "" Then
                                Call SetpageUchiwake7(strWksnme)
                            Else
                            '8ページ
                            If Sheets(strWksnme).Cells(446, 1) = "" Then
                                Call SetpageUchiwake8(strWksnme)
                            Else
                            '9ページ
                            If Sheets(strWksnme).Cells(501, 1) = "" Then
                                Call SetpageUchiwake9(strWksnme)
                            Else
                                MsgBox ("ページ想定外" & strWksnme)
                            End If
                            End If
                            End If
                            End If
                            End If
                            

                        End If
                    End If
                End If
            End If
        End If
    Next
    '***OFFにする
    Application.DisplayAlerts = True

End Sub
''*************************************************************************************************
''シートを個別にファイルとして出力する
''*************************************************************************************************
'Public Sub OutputSheetFile()
'    Dim intIdx As Integer       '処理用インデックス
'    Dim intWksCnt As Integer    '処理用カウンタ
'    Dim objWks As Object        'シート作成用オブジェクト
'    Dim strWbkNme As String     'Excelワークブック名(拡張子含まず)
'    Dim strWbkDir As String     'Excelワークブック保存場所
'    Dim strWksnme As String     'シート名
'
'    'Excelワークブックの情報取得
'    strWbkDir = Application.ActiveWorkbook.Path
'    strWbkNme = Application.ActiveWorkbook.Name
'    If Right(strWbkNme, Len(".xls")) = ".xls" Then
'        strWbkNme = Left(strWbkNme, Len(strWbkNme) - Len(".xls"))
'    End If
'
'    'シート数取得
'    intWksCnt = Excel.ActiveWorkbook.Worksheets.Count
'    '***OFFにする
'    Application.DisplayAlerts = False
'
'    'シート数ループ
'    For intIdx = 1 To intWksCnt
'        '対象シート名取得
'        strWksnme = Worksheets(intIdx).Name
'        If Right(strWksnme, 2) = "表紙" And strWksnme <> "表紙" Then
'            'シートのコピー
'            Worksheets(intIdx).Copy
'            'ファイル保存
'            ActiveWorkbook.SaveAs Filename:= _
'                strWbkDir & "\" & strWksnme & ".xls", _
'                FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
'                ReadOnlyRecommended:=False, CreateBackup:=False
'                Workbooks("ニッカホーム請求書作成.xlsm").Sheets(Left(strWksnme, 8) & "内訳").Copy Before:=Sheets(1)
'            ActiveWorkbook.Save
'
'            'Sheets(Array(Left(strWksnme, 8) & "表紙", Left(strWksnme, 8) & "内訳")).Select
'            Sheets(Array(Left(strWksnme, 8) & "内訳", Left(strWksnme, 8) & "表紙")).Select
'            'Sheets(Left(strWksNme, 8) & "表紙").Activate
'            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'                "C:\Users\KATOTO\Documents\メール送信\" & Left(strWksnme, 8) & ".pdf", Quality:=xlQualityStandard, _
'                IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
'            ActiveWindow.Close
'        End If
'    Next
'    '***OFFにする
'    Application.DisplayAlerts = True
'
'End Sub
'*************************************************************************************************
'シートを個別にファイルとして出力する
'*************************************************************************************************
Public Sub OutputSheetFile()
    Dim intIdx As Integer       '処理用インデックス
    Dim intWksCnt As Integer    '処理用カウンタ
    Dim objWks As Object        'シート作成用オブジェクト
    Dim strWbkNme As String     'Excelワークブック名(拡張子含まず)
    Dim strWbkDir As String     'Excelワークブック保存場所
    Dim strWksnme As String     'シート名
    Dim outFolder As String
    Dim outFolder2 As String

    'Excelワークブックの情報取得
    strWbkDir = Application.ActiveWorkbook.Path
    strWbkNme = Application.ActiveWorkbook.Name
    If Right(strWbkNme, Len(".xls")) = ".xls" Then
        strWbkNme = Left(strWbkNme, Len(strWbkNme) - Len(".xls"))
    End If

    'outFolder = "\\hob1sv07ap\ﾊﾟﾛﾏ共有$\会議用\ニッカホーム請求書\大問屋関東\"
    'outFolder1 = "\\hob1sv07ap\ﾊﾟﾛﾏ共有$\会議用\ニッカホーム請求書\ニッカホーム関東\"
    'outFolder2 = "\\hob1sv07ap\ﾊﾟﾛﾏ共有$\会議用\ニッカホーム請求書\その他\"
    'outFolder = "C:\Users\KATOTO\Documents\ニッカホーム\テスト\"
    'outFolder2 = "C:\Users\KATOTO\Documents\ニッカホーム\テスト\"
    outFolder = "\\HOB1SV03FS\販売⇔全国共有\■債権管理室■\■ニッカホーム請求書\大問屋関東\"
    outFolder1 = "\\HOB1SV03FS\販売⇔全国共有\■債権管理室■\■ニッカホーム請求書\ニッカホーム関東\"
    outFolder2 = "\\HOB1SV03FS\販売⇔全国共有\■債権管理室■\■ニッカホーム請求書\その他\"
    'outFolder = "C:\Users\KATOTO\Documents\ニッカホーム\テスト\大問屋関東\"
    'outFolder1 = "C:\Users\KATOTO\Documents\ニッカホーム\テスト\ニッカホーム関東\"
    'outFolder2 = "C:\Users\KATOTO\Documents\ニッカホーム\テスト\その他\"
    'シート数取得
    intWksCnt = Excel.ActiveWorkbook.Worksheets.Count
    '***OFFにする
    Application.DisplayAlerts = False

    'シート数ループ
    For intIdx = 1 To intWksCnt
        '対象シート名取得
        strWksnme = Worksheets(intIdx).Name
        If Right(strWksnme, 2) = "内訳" And strWksnme <> "内訳" Then
            'シートのコピー
            Worksheets(intIdx).Copy
            If Left(Workbooks("ニッカホーム請求書作成.xlsm").Sheets(Left(strWksnme, 8) & "表紙").Cells(6, 1), 3) = "大問屋" And Left(strWksnme, 2) = "03" Then
                'ファイル保存
                ActiveWorkbook.SaveAs Filename:= _
                    outFolder & strWksnme & ".xls", _
                    FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                    ReadOnlyRecommended:=False, CreateBackup:=False
                    Workbooks("ニッカホーム請求書作成.xlsm").Sheets(Left(strWksnme, 8) & "表紙").Copy Before:=Sheets(1)
                ActiveWorkbook.Save
    
                Sheets(Array(Left(strWksnme, 8) & "表紙", Left(strWksnme, 8) & "内訳")).Select
                'Sheets(Array(Left(strWksnme, 8) & "内訳", Left(strWksnme, 8) & "表紙")).Select
                'Sheets(Left(strWksNme, 8) & "表紙").Activate
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                    outFolder & Left(strWksnme, 8) & ".pdf", Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=Ture, OpenAfterPublish:=False
                ActiveWindow.Close
            Else
                If (Left(Workbooks("ニッカホーム請求書作成.xlsm").Sheets(Left(strWksnme, 8) & "表紙").Cells(6, 1), 6) = "ニッカホーム" Or _
                Left(Workbooks("ニッカホーム請求書作成.xlsm").Sheets(Left(strWksnme, 8) & "表紙").Cells(6, 1), 11) = "株式会社 ニッカホーム") And Left(strWksnme, 2) = "03" Then
                    'ファイル保存
                    ActiveWorkbook.SaveAs Filename:= _
                        outFolder1 & strWksnme & ".xls", _
                        FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                        ReadOnlyRecommended:=False, CreateBackup:=False
                        Workbooks("ニッカホーム請求書作成.xlsm").Sheets(Left(strWksnme, 8) & "表紙").Copy Before:=Sheets(1)
                    ActiveWorkbook.Save
        
                    Sheets(Array(Left(strWksnme, 8) & "表紙", Left(strWksnme, 8) & "内訳")).Select
                    'Sheets(Array(Left(strWksnme, 8) & "内訳", Left(strWksnme, 8) & "表紙")).Select
                    'Sheets(Left(strWksNme, 8) & "表紙").Activate
                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                        outFolder1 & Left(strWksnme, 8) & ".pdf", Quality:=xlQualityStandard, _
                        IncludeDocProperties:=True, IgnorePrintAreas:=Ture, OpenAfterPublish:=False
                    ActiveWindow.Close
                Else
                    'ファイル保存
                    ActiveWorkbook.SaveAs Filename:= _
                        outFolder2 & strWksnme & ".xls", _
                        FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                        ReadOnlyRecommended:=False, CreateBackup:=False
                        Workbooks("ニッカホーム請求書作成.xlsm").Sheets(Left(strWksnme, 8) & "表紙").Copy Before:=Sheets(1)
                    ActiveWorkbook.Save
        
                    Sheets(Array(Left(strWksnme, 8) & "表紙", Left(strWksnme, 8) & "内訳")).Select
                    'Sheets(Array(Left(strWksnme, 8) & "内訳", Left(strWksnme, 8) & "表紙")).Select
                    'Sheets(Left(strWksNme, 8) & "表紙").Activate
                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                        outFolder2 & Left(strWksnme, 8) & ".pdf", Quality:=xlQualityStandard, _
                        IncludeDocProperties:=True, IgnorePrintAreas:=Ture, OpenAfterPublish:=False
                    ActiveWindow.Close
                End If
            End If
        End If
    Next
    '***OFFにする
    Application.DisplayAlerts = True

End Sub
'*************************************************************************************************
'請求データを読み込む
'*************************************************************************************************
Public Sub makeNikkaHomeSeikyuData()
        
    Dim csvData As Variant
    Dim FSO As Object
    Dim intLine As Integer
    Dim intRow As Integer
    Dim i As Integer
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    intLine = 2
    
    Set csvfile = FSO.Opentextfile("C:\Users\KATOTO\Documents\ニッカホーム\NOHINPCX.CSV")
    'Set csvfile = FSO.Opentextfile("E:\NOHINPCX.CSV")
    With csvfile
        Do Until .atendofstream
            csvData = Split(Replace(.readline, """", ""), ",")
            intRow = 1
            For i = 0 To UBound(csvData)
                Sheets("データ").Cells(intLine, intRow) = csvData(i)
                intRow = intRow + 1
            Next i
            intLine = intLine + 1
        Loop
        .Close
    End With
    
    If Sheets("メニュー").Cells(1, 1) = "値引" Then
        
        '***値引データ
        Set csvfile = FSO.Opentextfile("E:\NOHINPCX.CSV")
        With csvfile
            Do Until .atendofstream
                csvData = Split(Replace(.readline, """", ""), ",")
                intRow = 1
                For i = 0 To UBound(csvData)
                    Sheets("データ").Cells(intLine, intRow) = csvData(i)
                    intRow = intRow + 1
                Next i
                intLine = intLine + 1
            Loop
            .Close
        End With
    End If
    TUDUKI_LINE = intLine
End Sub
'*************************************************************************************************
'請求データを読み込む
'2017.12.1 名古屋の要望で0648MA11,12,13を0648MA10にまとめても、各明細は出す
'*************************************************************************************************
Public Sub makeNikkaHomeSeikyuData_0648MA1X()
        
    Dim csvData As Variant
    Dim FSO As Object
    Dim intLine As Integer
    Dim intRow As Integer
    Dim i As Integer
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    intLine = TUDUKI_LINE
    
    Set csvfile = FSO.Opentextfile("C:\Users\KATOTO\Documents\ニッカホーム\NOHINPCX.CSV")
    With csvfile
        Do Until .atendofstream
            csvData = Split(Replace(.readline, """", ""), ",")
            intRow = 1
            If csvData(4) = "0648MA11" Or csvData(4) = "0648MA12" Or csvData(4) = "0648MA14" Or _
               csvData(4) = "0848ME21" Or csvData(4) = "0848ME22" Or csvData(4) = "0848ME23" Or csvData(4) = "0848ME24" Then
                For i = 0 To UBound(csvData)
                    Sheets("データ").Cells(intLine, intRow) = csvData(i)
                    intRow = intRow + 1
                Next i
                intLine = intLine + 1
            End If
        Loop
        .Close
    End With
    
    If Sheets("メニュー").Cells(1, 1) = "値引" Then
        
        '***値引データ
        Set csvfile = FSO.Opentextfile("E:\NOHINPCX.CSV")
        With csvfile
            Do Until .atendofstream
                csvData = Split(Replace(.readline, """", ""), ",")
                intRow = 1
                If csvData(4) = "0648MA11" Or csvData(4) = "0648MA12" Or csvData(4) = "0648MA14" Then
                    For i = 0 To UBound(csvData)
                        Sheets("データ").Cells(intLine, intRow) = csvData(i)
                        intRow = intRow + 1
                    Next i
                    intLine = intLine + 1
                End If
            Loop
            .Close
        End With
    End If
    TUDUKI_LINE = intLine
End Sub

'*************************************************************************************************
'読み込んだデータを
'*************************************************************************************************
Sub clearData()
'
    Sheets("データ").Select
    With Sheets("データ")
        Rows("2:2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.Delete Shift:=xlUp
    End With
    
End Sub
Sub CopyHyosi(ByVal CopySheetname As String, ByVal PasteSheetname As String)

    Sheets(CopySheetname).Select
    Cells.Select
    Cells.Copy
    Sheets(PasteSheetname).Select
    Sheets(PasteSheetname).Cells.Paste
End Sub
Sub Macro3()
'
' Macro3 Macro
'

'
    Cells.Select
End Sub

Sub Macro5()
'
' Macro5 Macro
'

'
    Sheets("表紙").Select
    Sheets("表紙").Copy After:=Sheets(5)
End Sub
Sub Macro6()
'
' Macro6 Macro
'

'
    Sheets("1148M240内訳").Select
    Sheets("1148M240内訳").Copy Before:=Workbooks("20160229QWE作成.xlsx").Sheets(1)
End Sub
Sub sortData()
'
' Macro1 Macro
'

'　　得意先、日付、注番号
    ActiveWorkbook.Worksheets("データ").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("データ").Sort.SortFields.Add Key:=Range("E2:E2500"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("データ").Sort.SortFields.Add Key:=Range("U2:U2500"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("データ").Sort.SortFields.Add Key:=Range("T2:T2500"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("データ").Sort.SortFields.Add Key:=Range("Q2:Q2500"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("データ").Sort
        .SetRange Range("A1:BC2500")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'****************************************************************************
'***配布先シートに貼り付けたあとに、関数を入れていく
'****************************************************************************
Private Sub chgdata(ByVal sheetName As String)
    Dim iniLine As Integer
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim i As Integer
    
    intLine = 1
    
    With Sheets(sheetName)
    
        Do While .Cells(intLine, 5) <> ""
            If Left(.Cells(intLine, 5), 7) = "0648M02" Then
                Debug.Print .Cells(intLine, 5) & " "; intLine
                .Cells(intLine, 5) = "0648M020"
            End If
            '***20171201 まとめ依頼
            If Left(.Cells(intLine, 5), 8) = "0648MA11" Or _
               Left(.Cells(intLine, 5), 8) = "0648MA15" Or _
               Left(.Cells(intLine, 5), 8) = "0648M011" Or _
               Left(.Cells(intLine, 5), 8) = "0648MA12" Or _
               Left(.Cells(intLine, 5), 8) = "0648MA14" Then
                Debug.Print .Cells(intLine, 5) & " "; intLine
                .Cells(intLine, 5) = "0648MA10"
            End If
            '***20181101 まとめ依頼
            If Left(.Cells(intLine, 5), 8) = "0848ME20" Or _
               Left(.Cells(intLine, 5), 8) = "0848ME21" Or _
               Left(.Cells(intLine, 5), 8) = "0848ME22" Or _
               Left(.Cells(intLine, 5), 8) = "0848ME23" Or _
               Left(.Cells(intLine, 5), 8) = "0848ME24" Then
                Debug.Print .Cells(intLine, 5) & " "; intLine
                .Cells(intLine, 5) = "0848ME20"
                .Cells(intLine, 6) = "エヌステージ(株)西日本"
            End If
            intLine = intLine + 1
        Loop
    End With

End Sub

'****************************************************************************
'***配布先シートに貼り付けたあとに、関数を入れていく
'****************************************************************************

Sub Setpage1(ByVal strWksnme As String)
'
' Macro4 Macro
'

'
     Sheets(strWksnme).Activate
'    Range("A37:U38").Select
'    Selection.Copy
'    Range("A37:U38").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'
'    Range("A10:H13").Select
'    Selection.Copy
'    Range("A10:H13").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'    Rows("39:90000").Select
'    Range(Selection, Selection.End(xlDown)).Select
'
'    Selection.Delete Shift:=xlUp
    
    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$U$37"
End Sub

'****************************************************************************
'***配布先シートに貼り付けたあとに、関数を入れていく
'****************************************************************************

Sub Setpage2(ByVal strWksnme As String)
'
' Macro4 Macro
'

'
     Sheets(strWksnme).Activate
'    Range("A37:U38").Select
'    Selection.Copy
'    Range("A37:U38").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'
'    Range("A10:H13").Select
'    Selection.Copy
'    Range("A10:H13").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'    Rows("39:90000").Select
'    Range(Selection, Selection.End(xlDown)).Select
'
'    Selection.Delete Shift:=xlUp
    
    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$U$72"
End Sub

'****************************************************************************
'***配布先シートに貼り付けたあとに、関数を入れていく
'****************************************************************************

Sub Setpage3(ByVal strWksnme As String)
'
' Macro4 Macro
'

'
     Sheets(strWksnme).Activate
    
    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$U$107"
End Sub

'****************************************************************************
'***配布先シートに貼り付けたあとに、関数を入れていく
'****************************************************************************

Sub Setpage4(ByVal strWksnme As String)
'
' Macro4 Macro
'

'
     Sheets(strWksnme).Activate
'    Range("A37:U38").Select
'    Selection.Copy
'    Range("A37:U38").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'
'    Range("A10:H13").Select
'    Selection.Copy
'    Range("A10:H13").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'    Rows("39:90000").Select
'    Range(Selection, Selection.End(xlDown)).Select
'
'    Selection.Delete Shift:=xlUp
    
    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$U$141"
End Sub


'****************************************************************************
'***配布先シートに貼り付けたあとに、関数を入れていく
'****************************************************************************

Sub SetpageUchiwake1(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$55"
End Sub

Sub SetpageUchiwake2(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$110"
End Sub

Sub SetpageUchiwake3(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$165"
End Sub


Sub SetpageUchiwake4(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$220"
End Sub


Sub SetpageUchiwake5(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$275"
End Sub

Sub SetpageUchiwake6(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$330"
End Sub

Sub SetpageUchiwake7(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$385"
End Sub

Sub SetpageUchiwake8(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$440"
End Sub

Sub SetpageUchiwake9(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$495"
End Sub
