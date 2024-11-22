Attribute VB_Name = "Module1"
Sub dataTotalling()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet

    Set ws1 = ThisWorkbook.Sheets("�\��t���V�[�g")
    Set ws2 = ThisWorkbook.Sheets("�W�v�V�[�g")
    '�e�[�u���I�u�W�F�N�g�擾
    Set tbl1 = ws1.ListObjects(1)
    Set tbl2 = ws2.ListObjects("�W�v�e�[�u��")
    
    ' �A�z�z���`
    Dim dataDic As Object
    Dim dataKey As Variant
    Set dataDic = CreateObject("Scripting.Dictionary")
    
    ' �H���ԍ��ƌ���䗦���i�[����ϐ����`
    Dim kojiNum1 As ListColumn
    Dim kojiNum2 As ListColumn
    Dim kojiNum3 As ListColumn
    Dim rate1 As ListColumn
    Dim rate2 As ListColumn
    Dim rate3 As ListColumn
    Dim totalKojiNum As ListColumn
    Dim totalRate As ListColumn
    
    Set kojiNum1 = tbl1.ListColumns("�@�H���ԍ�")
    Set rate1 = tbl1.ListColumns("�@����䗦")
    
    Set kojiNum2 = tbl1.ListColumns("�A�H���ԍ�")
    Set rate2 = tbl1.ListColumns("�A����䗦")
    
    Set kojiNum3 = tbl1.ListColumns("�B�H���ԍ�")
    Set rate3 = tbl1.ListColumns("�B����䗦")
    
    Set totalKojiNum = tbl2.ListColumns("�H���ԍ�")
    Set totalRate = tbl2.ListColumns("�l�H")
    Set totalCost = tbl2.ListColumns("�l����")
    
    ' �e�[�u���̍s���J��Ԃ��@�A�z�z��Ɋi�[
    For Each ListRow In tbl1.ListRows
        ' ���ݗ�̇@�A�B�̍H���ԍ��Ɣ䗦���i�[
        kNumValue1 = ListRow.Range.Cells(1, kojiNum1.Index).Value
        rateValue1 = ListRow.Range.Cells(1, rate1.Index).Value
        
        kNumValue2 = ListRow.Range.Cells(1, kojiNum2.Index).Value
        rateValue2 = ListRow.Range.Cells(1, rate2.Index).Value
        
        kNumValue3 = ListRow.Range.Cells(1, kojiNum3.Index).Value
        rateValue3 = ListRow.Range.Cells(1, rate3.Index).Value
       
        ' �@�ɂ���
        ' �z��ɑ��݂��Ȃ��Ȃ�V�K�ǉ��@���݂���Ȃ�䗦�ɉ��Z
        If Not dataDic.Exists(kNumValue1) Then
            dataDic.Add kNumValue1, rateValue1
        Else
            tmpRateValue = dataDic(kNumValue1)
            dataDic(kNumValue1) = tmpRateValue + rateValue1
        End If
        
        ' �A�B�H���ԍ��̔���
        ' �A�̋�`�F�b�N
        If Not kNumValue2 = "" Then
            If Not dataDic.Exists(kNumValue2) Then
                dataDic.Add kNumValue2, rateValue2
            Else
                tmpRateValue = dataDic(kNumValue2)
                dataDic(kNumValue2) = tmpRateValue + rateValue2
            End If
        End If
        
        ' �B�̋�`�F�b�N
        If Not kNumValue3 = "" Then
            If Not dataDic.Exists(kNumValue3) Then
                dataDic.Add kNumValue3, rateValue3
            Else
                tmpRateValue = dataDic(kNumValue3)
                dataDic(kNumValue3) = tmpRateValue + rateValue3
            End If
        End If
    Next ListRow
    
    ' �W�v�V�[�g�̃e�[�u��������
    If Not tbl2.DataBodyRange Is Nothing Then
        tbl2.DataBodyRange.Delete
    End If
    
    ' �W�v�V�[�g�̃e�[�u���ɓ]�L���Ă���
    For Each dataKey In dataDic.Keys
        tbl2.ListRows.Add
        totalKojiNum.DataBodyRange(tbl2.ListRows.Count).Value = dataKey
        totalRate.DataBodyRange(tbl2.ListRows.Count).Value = dataDic.Item(dataKey)
        totalCost.DataBodyRange(tbl2.ListRows.Count).Value = dataDic.Item(dataKey) * 20000
    Next
    
    ' �l������~���\�[�g
    tbl2.Range.Sort key1:=totalCost.Range, order1:=xlDescending, Header:=xlYes
    
End Sub

Sub clearWs1()
    Dim ws1 As Worksheet

    Set ws1 = ThisWorkbook.Sheets("�\��t���V�[�g")
    Set tbl1 = ws1.ListObjects(1)
   
    If Not tbl1.DataBodyRange Is Nothing Then
        tbl1.DataBodyRange.Delete
    End If
    
End Sub

Sub clearWs2()
    Dim ws2 As Worksheet

    Set ws2 = ThisWorkbook.Sheets("�W�v�V�[�g")
    Set tbl2 = ws2.ListObjects("�W�v�e�[�u��")
    
    If Not tbl2.DataBodyRange Is Nothing Then
        tbl2.DataBodyRange.Delete
    End If
    
End Sub
