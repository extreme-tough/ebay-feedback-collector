Attribute VB_Name = "Module1"
Public Type FeedBack
    ItemName As String
    ItemPrice As String
    ItemDate As String
    ItemNumber As String
End Type

Public Sub QuickSort(ByRef MyArray() As FeedBack, ByVal lngLBound As Long, ByVal lngUBound As Long)
    Dim fbPivot As FeedBack
    Dim k As Long
    Dim p As Long
    Dim fbTemp As FeedBack
        
    
    If lngLBound >= lngUBound Then
        Exit Sub
    End If
    
    k = lngLBound + 1
    
    If lngUBound = k Then
        If MyArray(lngLBound).ItemName > MyArray(lngUBound).ItemName Then
            'swap MyArray(lngLBound) and MyArray(lngUBound)
            fbTemp = MyArray(lngLBound)
            MyArray(lngLBound) = MyArray(lngUBound)
            MyArray(lngUBound) = fbTemp
        End If
        Exit Sub
    End If
    
    fbPivot = MyArray(lngLBound)
    p = lngUBound
    
    Do Until (MyArray(k).ItemName > fbPivot.ItemName) Or (k >= lngUBound)
        k = k + 1
    Loop
    
    Do Until MyArray(p).ItemName <= fbPivot.ItemName
        p = p - 1
    Loop
    
    Do While k < p
        'swap MyArray(k) and MyArray(p)
        fbTemp = MyArray(k)
        MyArray(k) = MyArray(p)
        MyArray(p) = fbTemp
        
        Do
            k = k + 1
        Loop Until MyArray(k).ItemName > fbPivot.ItemName
        
        Do
            p = p - 1
        Loop Until MyArray(p).ItemName <= fbPivot.ItemName
    Loop
    
    'swap MyArray(p) and MyArray(lngLBound)
    fbTemp = MyArray(p)
    MyArray(p) = MyArray(lngLBound)
    MyArray(lngLBound) = fbTemp
    
    QuickSort MyArray, lngLBound, p - 1
    QuickSort MyArray, p + 1, lngUBound
End Sub

Public Sub WriteRawData(MyArray() As FeedBack, ByVal SellerName As String)
    Dim i  As Integer
    Dim oExcel As New Excel.Application
    Dim oWB As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    
    Dim TargetFile As String
    
    TargetFile = App.Path + "\output\" + SellerName + ".xls"
    On Error Resume Next
    Kill TargetFile
    On Error GoTo 0
    
    FileCopy App.Path + "\template.xls", TargetFile
    
    Set oWB = oExcel.Workbooks.Open(TargetFile)
    Set oSheet = oWB.Sheets(1)

    For i = LBound(MyArray) To UBound(MyArray)
        oSheet.Cells(i + 1, 1).Value = MyArray(i).ItemName + "(" + MyArray(i).ItemNumber + ")"
        oSheet.Cells(i + 1, 2).Value = Replace(MyArray(i).ItemPrice, "US", "")
        oSheet.Cells(i + 1, 3).Value = MyArray(i).ItemDate
        oSheet.Cells(i + 1, 4).Value = SellerName
    Next
    oWB.Save
    oExcel.Quit
    
End Sub

Public Sub WriteSortedData(MyArray() As FeedBack, ByVal SellerName As String)
    Dim i  As Integer
    Dim oExcel As New Excel.Application
    Dim oWB As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim lastItemName As String
    Dim DupCount As Long
    Dim oRng As Excel.Range


    Dim TargetFile As String
    
    TargetFile = App.Path + "\output\" + SellerName + ".xls"

    
    Set oWB = oExcel.Workbooks.Open(TargetFile)
    Set oSheet = oWB.Sheets(2)
    lastItemName = ""
    DupCount = 1
    
    For i = LBound(MyArray) To UBound(MyArray)
        If MyArray(i).ItemName = lastItemName Then
            DupCount = DupCount + 1
        Else
            DupCount = 1
        End If
        
        If DupCount > 1 Then
            If i < UBound(MyArray) Then
                If MyArray(i).ItemName <> MyArray(i + 1).ItemName Then
                    oSheet.Cells(i + 1, 1).Value = DupCount
                End If
            Else
                oSheet.Cells(i + 1, 1).Value = DupCount
            End If
            
            Set oRng = oSheet.Rows(i + 1)
            
            With oRng.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With

        End If
        
        oSheet.Cells(i + 1, 2).Value = MyArray(i).ItemName + "(" + MyArray(i).ItemNumber + ")"
        oSheet.Cells(i + 1, 3).Value = Replace(MyArray(i).ItemPrice, "US", "")
        oSheet.Cells(i + 1, 4).Value = MyArray(i).ItemDate
        oSheet.Cells(i + 1, 5).Value = SellerName
        lastItemName = MyArray(i).ItemName
    Next
    
    oWB.Save
    oExcel.Quit
    
End Sub
Public Function GetSellerName(ByVal URL As String)
    Dim nPos1 As Integer
    nPos1 = InStr(1, URL, "userid=", VbCompareMethod.vbTextCompare)
    If nPos1 < 0 Then
        GetSellerName = ""
        Exit Function
    End If
    
    npos2 = InStr(nPos1, URL, "&")
    If npos2 < nPos1 Or npos2 < 0 Then
        GetSellerName = ""
        Exit Function
    End If
        
    GetSellerName = Mid$(URL, nPos1 + 7, npos2 - nPos1 - 7)
End Function

