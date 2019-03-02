Attribute VB_Name = "Module1"
Option Explicit

Const mc_MaxHeadingLevel As Long = 3

Dim m_WS As Worksheet
Dim m_DataRange As Range

Sub createIndex()
  Set m_WS = ActiveSheet
  Set m_DataRange = m_WS.Range("A3")
  
  Dim Filepath As String
  With Application.FileDialog(msoFileDialogOpen)
    .AllowMultiSelect = False
    .Filters.Clear
    .Filters.Add "MDファイル(*.md)", "*.md", 1
    .InitialFileName = ThisWorkbook.Path & "\"
    .Show
    
    If .SelectedItems.Count = 0 Then
      End
    End If
    
    Filepath = .SelectedItems(1)
  End With
 
  Dim ContentStr As String
  ContentStr = ReadCSV.ReadFileToString(Filepath)
  
  m_WS.UsedRange.Clear
  Call createIndexImpl(ContentStr)

  Call FormatLevelStyle(1, "", "章 ", True)
  Call FormatLevelStyle(2, "", " ", True)
  
  Call InputAndFormatTitle
  
End Sub

'mc_HeadingLevelを変えた場合にはやり直し
Private Sub InputAndFormatTitle()
  m_WS.Cells(2, 1) = "章"
  m_WS.Cells(2, 2) = "項"
  m_WS.Cells(2, 3) = "見出し"
  m_WS.Cells(2, 4) = "備考・検討事項"

  m_WS.Cells(2, 1).Resize(1, mc_MaxHeadingLevel + 1).Interior.ColorIndex = 15
  
'2列目から列幅自動調整　※1、2列目はそのまま
  m_WS.Range(m_WS.Columns(3), m_WS.Columns(mc_MaxHeadingLevel + 1)).AutoFit
  
End Sub

Private Sub FormatLevelStyle(Optional Level As Long, Optional Prefix As String = "", Optional Postfix As String = "", Optional FontBold = False)
  Dim Count As Long
  Count = 1
  
  Dim Row As Long
  
  With m_WS
    For Row = m_DataRange.Row To m_WS.Cells(Rows.Count, 1).End(xlUp).Row
      If .Cells(Row, Level) <> "" Then
        .Cells(Row, Level).Value = Prefix & Count & Postfix & .Cells(Row, Level).Value
        .Cells(Row, Level).Font.Bold = FontBold
        Count = Count + 1
      End If
    Next
  End With
End Sub


Private Sub createIndexImpl(ContentStr As String)

  Dim ContentLines As Variant
  ContentLines = Split(ContentStr, vbCrLf)
  
  Dim Row As Long
  Row = 1
  
  Dim HeadingLevel As Long
  
'仮インデックス作成
  Dim Line As Variant
  For Each Line In ContentLines
    HeadingLevel = DetectHeadingLevel(CStr(Line), mc_MaxHeadingLevel)
    
    'すでに同階層の見出しが転記済の場合は、改行
    If HeadingLevel >= 1 And HeadingLevel <= mc_MaxHeadingLevel Then
      If WorksheetFunction.CountA(m_DataRange.Cells(Row, HeadingLevel).Resize(1, 6 - HeadingLevel + 1)) >= 1 Then
        Row = Row + 1
      End If
     
      m_DataRange.Cells(Row, HeadingLevel).Value = Mid(Line, HeadingLevel + 2, Len(Line)) 'Len→Strに変えると落ちる
      
    End If
    
    '#見出しの後は無条件で改行
    If HeadingLevel = 1 Then
      Row = Row + 1
    End If
    
    '#＃見出しの後は無条件で改行
    If HeadingLevel = 2 Then
      Row = Row + 1
    End If
    
  Next
End Sub


Private Function DetectHeadingLevel(Line As String, Optional MaxLevel As Long = 3) As Long
  Dim Count As Long
  For Count = 6 To 1 Step -1
    If Left(Line, Count) = String(Count, "#") Then
      DetectHeadingLevel = Count
      Exit Function
    End If
  Next
  
  DetectHeadingLevel = 0
End Function
