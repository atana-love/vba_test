Attribute VB_Name = "Module1"
Sub dataTotalling()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet

    Set ws1 = ThisWorkbook.Sheets("貼り付けシート")
    Set ws2 = ThisWorkbook.Sheets("集計シート")
    'テーブルオブジェクト取得
    Set tbl1 = ws1.ListObjects(1)
    Set tbl2 = ws2.ListObjects("集計テーブル")
    
    ' 連想配列定義
    Dim dataDic As Object
    Dim dataKey As Variant
    Set dataDic = CreateObject("Scripting.Dictionary")
    
    ' 工事番号と現場比率を格納する変数を定義
    Dim kojiNum1 As ListColumn
    Dim kojiNum2 As ListColumn
    Dim kojiNum3 As ListColumn
    Dim rate1 As ListColumn
    Dim rate2 As ListColumn
    Dim rate3 As ListColumn
    Dim totalKojiNum As ListColumn
    Dim totalRate As ListColumn
    
    Set kojiNum1 = tbl1.ListColumns("①工事番号")
    Set rate1 = tbl1.ListColumns("①現場比率")
    
    Set kojiNum2 = tbl1.ListColumns("②工事番号")
    Set rate2 = tbl1.ListColumns("②現場比率")
    
    Set kojiNum3 = tbl1.ListColumns("③工事番号")
    Set rate3 = tbl1.ListColumns("③現場比率")
    
    Set totalKojiNum = tbl2.ListColumns("工事番号")
    Set totalRate = tbl2.ListColumns("人工")
    Set totalCost = tbl2.ListColumns("人件費")
    
    ' テーブルの行分繰り返し　連想配列に格納
    For Each ListRow In tbl1.ListRows
        ' 現在列の①②③の工事番号と比率を格納
        kNumValue1 = ListRow.Range.Cells(1, kojiNum1.Index).Value
        rateValue1 = ListRow.Range.Cells(1, rate1.Index).Value
        
        kNumValue2 = ListRow.Range.Cells(1, kojiNum2.Index).Value
        rateValue2 = ListRow.Range.Cells(1, rate2.Index).Value
        
        kNumValue3 = ListRow.Range.Cells(1, kojiNum3.Index).Value
        rateValue3 = ListRow.Range.Cells(1, rate3.Index).Value
       
        ' ①について
        ' 配列に存在しないなら新規追加　存在するなら比率に加算
        If Not dataDic.Exists(kNumValue1) Then
            dataDic.Add kNumValue1, rateValue1
        Else
            tmpRateValue = dataDic(kNumValue1)
            dataDic(kNumValue1) = tmpRateValue + rateValue1
        End If
        
        ' ②③工事番号の判定
        ' ②の空チェック
        If Not kNumValue2 = "" Then
            If Not dataDic.Exists(kNumValue2) Then
                dataDic.Add kNumValue2, rateValue2
            Else
                tmpRateValue = dataDic(kNumValue2)
                dataDic(kNumValue2) = tmpRateValue + rateValue2
            End If
        End If
        
        ' ③の空チェック
        If Not kNumValue3 = "" Then
            If Not dataDic.Exists(kNumValue3) Then
                dataDic.Add kNumValue3, rateValue3
            Else
                tmpRateValue = dataDic(kNumValue3)
                dataDic(kNumValue3) = tmpRateValue + rateValue3
            End If
        End If
    Next ListRow
    
    ' 集計シートのテーブル初期化
    If Not tbl2.DataBodyRange Is Nothing Then
        tbl2.DataBodyRange.Delete
    End If
    
    ' 集計シートのテーブルに転記していく
    For Each dataKey In dataDic.Keys
        tbl2.ListRows.Add
        totalKojiNum.DataBodyRange(tbl2.ListRows.Count).Value = dataKey
        totalRate.DataBodyRange(tbl2.ListRows.Count).Value = dataDic.Item(dataKey)
        totalCost.DataBodyRange(tbl2.ListRows.Count).Value = dataDic.Item(dataKey) * 20000
    Next
    
    ' 人件費を降順ソート
    tbl2.Range.Sort key1:=totalCost.Range, order1:=xlDescending, Header:=xlYes
    
End Sub

Sub clearWs1()
    Dim ws1 As Worksheet

    Set ws1 = ThisWorkbook.Sheets("貼り付けシート")
    Set tbl1 = ws1.ListObjects(1)
   
    If Not tbl1.DataBodyRange Is Nothing Then
        tbl1.DataBodyRange.Delete
    End If
    
End Sub

Sub clearWs2()
    Dim ws2 As Worksheet

    Set ws2 = ThisWorkbook.Sheets("集計シート")
    Set tbl2 = ws2.ListObjects("集計テーブル")
    
    If Not tbl2.DataBodyRange Is Nothing Then
        tbl2.DataBodyRange.Delete
    End If
    
End Sub
