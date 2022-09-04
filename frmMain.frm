VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "eBay Feedback Collector v1.0"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   1455
      Left            =   3720
      TabIndex        =   4
      Top             =   5520
      Width           =   735
      ExtentX         =   1296
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   10320
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Start"
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtURLList 
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   11415
   End
   Begin VB.Label lblSubStatus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   45
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "Enter seller feedback URLs here (One URL per Line)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
'http://feedback.ebay.com/ws/eBayISAPI.dll?ViewFeedback2&ftab=FeedbackAsSeller&userid=allbmwparts&iid=-1&de=off&items=200

Dim colFeedBacks() As FeedBack

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdProcess_Click()
    Dim arURLs() As String
    Dim i As Integer
    Dim SellerName As String
    
    If Trim(txtURLList.Text) = "" Then
        MsgBox "Please enter the URLs in the text box", vbInformation
        Exit Sub
    End If
    
    cmdProcess.Enabled = False
    
    
    
    arURLs = Split(txtURLList.Text, vbCrLf)
    For i = 0 To UBound(arURLs)
        If Trim$(arURLs(i)) <> "" Then
            lblStatus.Caption = "Processing URL " + Str(i + 1) + " of " + Str(UBound(arURLs) + 1)
            lblStatus.Refresh
            SellerName = GetSellerName(arURLs(i))
            If SellerName = "" Then
                MsgBox "Seller name not identifiable in URL - " + arURLs(i), vbCritical
                Exit Sub
            End If
            lblSubStatus.Caption = "Getting the feedback data..."
            lblSubStatus.Refresh
            'On Error Resume Next
            colFeedBacks = GetFeedBack(arURLs(i))
            If Err Then
                MsgBox "Could not load URL " + arURLs(i) + vbCrLf _
                        + "Error : " + Err.Description + ":" + Err.Source, vbCritical
            Else
                lblSubStatus.Caption = "Writing raw data..."
                lblSubStatus.Refresh
                WriteRawData colFeedBacks, SellerName
                lblSubStatus.Caption = "Sorting data..."
                lblSubStatus.Refresh
                QuickSort colFeedBacks, LBound(colFeedBacks), UBound(colFeedBacks)
                lblSubStatus.Caption = "Writing sorted data..."
                lblSubStatus.Refresh
                WriteSortedData colFeedBacks, SellerName
                lblSubStatus.Caption = ""
                lblSubStatus.Refresh
            End If
            On Error GoTo 0
        End If
    Next
    MsgBox "Files are copied to " + App.Path + "\output folder", vbInformation
    cmdProcess.Enabled = True
End Sub


Private Function GetFeedBack(sURL As String) As FeedBack()
    Dim Content As String
    Dim i As Integer
    Dim j As Integer
    
    Dim EHTML As Variant, TableElem As Variant, colElem As Variant, itemElems As Variant
    
    Dim FeedBack As String
    Dim FeedDate As String
    Dim ItemName As String
    Dim PartNumber As String
    Dim sPrice As String
    Dim Page As Integer
    Dim TableFound As Boolean
    Dim lastPageData As String
    Dim CollectedData() As FeedBack
    Dim RowsCollected As Integer
    Dim dt1 As Date
    Dim dt2 As Date
    
    Page = 1
    RowsCollected = 0
    
    lastPageData = ""
    
    Do While True
    
    
        dt1 = Now()
        WebBrowser.Navigate sURL + "&page=" + Trim$(Str(Page))
        
        Do While WebBrowser.ReadyState <> READYSTATE_COMPLETE
            dt2 = Now()
            If DateDiff("m", dt1, dt2) > 2 Then
                Err.Raise 100, "Timeout loading " + sURL
                Exit Function
            End If
            DoEvents
        Loop
        
        
        If WebBrowser.Document.body.innertext = lastPageData Then
            'Same page is loaded again
            GetFeedBack = CollectedData
            Exit Function
        End If
        
        lastPageData = WebBrowser.Document.body.innertext
        
        TableFound = False
        
        'For i = 0 To WebBrowser.Document.All.Length - 1
        For i = 0 To WebBrowser.Document.getElementsbytagname("TABLE").Length - 1
            Set EHTML = WebBrowser.Document.getElementsbytagname("TABLE")(i)
            If Not (EHTML Is Nothing) Then
                'If EHTML.TagName = "TABLE" Then
                    If InStr(1, EHTML.OuterHTML, "FbOuterYukon") > 0 Then
                        For j = 0 To EHTML.All.Length - 1
                            Set TableElem = EHTML.All.Item(j)
                            If TableElem.TagName = "TR" Then
                                Content = Trim(Replace(TableElem.OuterHTML, vbCrLf, ""))
                                If Left(Content, 4) = "<TR>" Then
                                    Set itemElems = TableElem.getElementsbytagname("TD")
                                    If InStr(1, Content, "<B>Follow-up</B>") > 0 Or _
                                    InStr(1, Content, "<B>Reply</B>") > 0 Or _
                                    InStr(1, Content, "Feedback was revised on") > 0 Then
                                        'Ignore if comment contains follow up data or reply
                                    Else
                                        If InStr(1, itemElems.Item(2).innertext, "Detailed item information is not available for the following") > 0 Then
                                            GetFeedBack = CollectedData
                                            Exit Function
                                        End If
                                        FeedDate = itemElems.Item(3).innertext
                                    
                                        If DateDiff("D", CDate(FeedDate), Now) > 90 Then
                                            GetFeedBack = CollectedData
                                            Exit Function
                                        End If
                                        RowsCollected = RowsCollected + 1
                                        ReDim Preserve CollectedData(1 To RowsCollected) As FeedBack
                                        CollectedData(RowsCollected).ItemDate = FeedDate
                                    End If
                                End If
                                If Left(Content, 14) = "<TR class=bot>" Then
                                    Set itemElems = TableElem.getElementsbytagname("TD")
                                    ItemName = Split(itemElems.Item(1).innertext, "(#")(0)
                                    If ItemName <> "--" Then
                                        CollectedData(RowsCollected).ItemName = ItemName
                                        PartNumber = Replace(Split(itemElems.Item(1).innertext, "(#")(1), ")", "")
                                        sPrice = itemElems.Item(2).innertext
                                        CollectedData(RowsCollected).ItemPrice = sPrice
                                        CollectedData(RowsCollected).ItemNumber = PartNumber
                                    Else
                                        RowsCollected = RowsCollected + 1
                                    End If
                                End If
                            End If
                        Next
                        Exit For
                    End If
                    TableFound = True
                'End If
            End If
        Next
        
        If Not TableFound Then
            GetFeedBack = CollectedData
            Exit Function
        End If
        
        Page = Page + 1
    Loop
End Function



