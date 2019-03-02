Attribute VB_Name = "ReadCSV"
Option Explicit

'NumberFormats ��Dictionary�iKey�F��ԍ��AValue�F���[�U��`�����j���w�肷��
'Charset�͏����ݒ�̑��AShift_JIS�Aeuc-jp�Autf-8�Autf-8n�Ȃǂ��ݒ�ł���

Const mc_AutoDetectString As String = "_autodetect_all"

Public Function ReadCSV(ByRef Filepath As String, _
    Optional ByRef OutputWorksheet As Worksheet, _
    Optional ByRef TextColumns As String = "", _
    Optional ByRef SkipColumns As String = "", _
    Optional ByRef ColumnNumberFormats As Object = Nothing, _
    Optional ByRef ReadHeader As Boolean = True, _
    Optional ByVal OutputRow As Long = 1, _
    Optional ByVal OutputColumn As Long = 1, _
    Optional ByRef Delimiter As String = ",", _
    Optional ByRef Charset As String = mc_AutoDetectString, _
    Optional ByRef Quote As String = """", _
    Optional ByRef LineEndingCode As String = mc_AutoDetectString, _
    Optional ByRef AutoFit As Boolean = True) As Worksheet

  Dim CSV As Variant
  CSV = ReadCSVToArray2D(Filepath, SkipColumns, ReadHeader, Delimiter, Charset, Quote, LineEndingCode)
  
  If OutputWorksheet Is Nothing Then
    Set OutputWorksheet = Worksheets.Add
  End If
  
  Call Array2DToWorksheetWithColumnFormat(OutputWorksheet, CSV, TextColumns, SkipColumns, ColumnNumberFormats, OutputRow, OutputColumn, AutoFit)
  
  Set ReadCSV = OutputWorksheet
End Function

Public Sub Array2DToWorksheetWithColumnFormat(ByRef OutputWorksheet As Worksheet, _
    ByRef Array2D As Variant, _
    Optional ByRef TextColumns As String = "", _
    Optional ByRef SkipColumns As String = "", _
    Optional ByRef ColumnNumberFormats As Object = Nothing, _
    Optional ByVal OutputRow As Long = 1, _
    Optional ByVal OutputColumn As Long = 1, _
    Optional ByRef AutoFit As Boolean = True)
  
  Dim ColumnNumberFormat As Variant
  ColumnNumberFormat = getColumnDataTypes(TextColumns, SkipColumns, ColumnNumberFormats, UBound(Array2D, 2) - LBound(Array2D, 2) + 1)
   
  Dim maxRow As Long
  maxRow = OutputRow + UBound(Array2D, 1) - LBound(Array2D, 1)
  
  Dim maxColumn As Long
  maxColumn = OutputColumn + UBound(Array2D, 2) - LBound(Array2D, 2)
  
  With OutputWorksheet
    Dim oCol As Long
    
    For oCol = 1 To UBound(Array2D, 2) - LBound(Array2D, 2) + 1
      .Range(.Cells(OutputRow, OutputColumn + oCol - 1), _
             .Cells(maxRow, OutputColumn + oCol - 1)).NumberFormatLocal = ColumnNumberFormat(oCol)
    Next
      
    .Range(.Cells(OutputRow, OutputColumn), .Cells(maxRow, maxColumn)).Value = Array2D
    
    If AutoFit Then
      .Range(.Columns(OutputColumn), .Columns(maxColumn)).AutoFit
    End If
  End With

End Sub


'NumberFormats��Dictionary�O��
Private Function getColumnDataTypes(TextColumns As String, SkipColumns As String, ColumnNumberFormats As Object, MaxColumnNumber As Long) As Variant
  
  Dim SkipColumnDict As Object
  Set SkipColumnDict = ParseColumnsSelectString(SkipColumns)
  
'��������A����SkipColumnDict���Ȃ����̂Ƃ��Ĕz����쐬
  Dim FormatData As Variant
  ReDim FormatData(1 To MaxColumnNumber + SkipColumnDict.Count)
  
  Dim C As Long
  For C = 1 To MaxColumnNumber + SkipColumnDict.Count
    FormatData(C) = "G/�W��"
  Next
  
  Dim Column As Variant
  
  Dim TextColumnDict As Object
  Set TextColumnDict = ParseColumnsSelectString(TextColumns)
  
  For Each Column In TextColumnDict
    FormatData(CLng(Column)) = "@"
  Next


'���ڍׂȕ\���`���w�肪����ꍇ�ɂ́A�㏑��
  If Not ColumnNumberFormats Is Nothing Then
    For Each Column In ColumnNumberFormats.Keys
      FormatData(Column) = ColumnNumberFormats(Column)
    Next
  End If
  
  
'Skip�����̕␳������
'Skip�Ώۂ̗񂪂���ꍇ�ɂ́A�����Ƃ̑Ή��������̂ŕ␳������
'���Ƃ��΁ACSV��4��ŁA�������u������,������,skip,�W���v�Ƃ����w��ɂȂ��Ă���ꍇ�B
'CSV��4��ڂ̏����Ƃ��Ďw�肵�Ă���l�i�W���j���A�V�[�g��3��ڂɐݒ肵�Ȃ��Ƃ����Ȃ�
   
  Dim Ret As Variant
  ReDim Ret(1 To MaxColumnNumber)
  
  Dim RetColumn As Long
  RetColumn = LBound(Ret)
  
  Dim FormatColumn As Long
  For FormatColumn = LBound(FormatData) To UBound(FormatData)
    If Not SkipColumnDict.Exists(CStr(FormatColumn)) Then
      Ret(RetColumn) = FormatData(FormatColumn)
      RetColumn = RetColumn + 1
    End If
  Next

Finally:
  getColumnDataTypes = Ret
End Function


Private Function ParseColumnsSelectString(ColumnsSelectString As String) As Object
  Dim Columns As Variant
  Columns = Split(ColumnsSelectString, ",")
  
  Dim Dict As Object
  Set Dict = CreateObject("Scripting.Dictionary")
  
  Dim Column As Variant
  For Each Column In Columns
    If Not Dict.Exists(Column) Then
      Dict(Column) = Column
    End If
  Next
  
  Set ParseColumnsSelectString = Dict
  
End Function


'LineEndingCode�ɂ́AvbCrLf vbCr vbLf�Ȃǂ��w�肷��
Public Function ReadCSVToArray2D(ByRef inputFilepath As String, _
    Optional ByRef SkipColumns As String = "", _
    Optional ByVal ReadHeader As Boolean = True, _
    Optional ByRef Delimiter As String = ",", _
    Optional ByRef Charset As String = mc_AutoDetectString, _
    Optional ByRef Quote As String = """", _
    Optional ByRef LineEndingCode As String = mc_AutoDetectString) As Variant
  
  Dim InputStr As String
  InputStr = ReadFileToString(inputFilepath, Charset)
  
  ReadCSVToArray2D = CSVStringToArray2D(InputStr, SkipColumns, ReadHeader, Delimiter, Quote, LineEndingCode)
  
End Function


'CSV�f�[�^���i�[�������������͂���2�����z���Ԃ�
'LineEndingCode�ɂ́AvbCrLf vbCr vbLf�Ȃǂ��w�肷��
Public Function CSVStringToArray2D(ByRef InputStr As String, _
    Optional ByRef SkipColumns As String = "", _
    Optional ByVal ReadHeader As Boolean = True, _
    Optional ByRef Delimiter As String = ",", _
    Optional ByRef Quote As String = """", _
    Optional ByRef LineEndingCode As String = mc_AutoDetectString) As Variant
    
  If Len(Quote) > 1 Then
    MsgBox "���p����1�����Ŏw�肵�Ă�������"
    End
  End If
  
  If LineEndingCode = mc_AutoDetectString Then
    LineEndingCode = DetectLineEndingCode(InputStr)
  End If
  
'�Ƃ肠�����A�s���̍ő�l���擾���邽�߁A�b��I�ɓ��̓f�[�^�����s�R�[�h���Ƃɔz��ɕ���
  Dim iLines() As String
  iLines = Split(InputStr, LineEndingCode)

  Dim oLinesCollection As Collection
  Set oLinesCollection = New Collection
   
  Dim Re As Object
  Set Re = getRegexPattern(Delimiter, Quote)
  
  Dim SkipColumnDict As Object
  Set SkipColumnDict = ParseColumnsSelectString(SkipColumns)
  
  
  Dim oLine As Variant  '1�����z�񂪓���
  Dim oLinesColumnMax As Long
  
  'iLines�̌��ݏ����Ώۍs
  Dim iY As Long
  iY = LBound(iLines)
  
  'CSV����͂������ʂ���������Collection�ioLine�j�Ɋi�[
  '�ő�񐔂�oLinesColumnMax�Ɋi�[
  Do While iY <= UBound(iLines)
    
    'getLine�̌Ăяo����AiY���ω�����i������j�̂Œ��ӁI
    Dim iLine As String
    iLine = GetLine(iLines, iY, LineEndingCode, Quote)
        
    Set oLine = ParseLine(iLine, Re, LineEndingCode, Quote, Delimiter, SkipColumnDict)
    oLinesCollection.Add oLine
    oLinesColumnMax = WorksheetFunction.Max(oLine.Count, oLinesColumnMax)
    
  Loop
  
  If Not ReadHeader Then
    oLinesCollection.Remove 1
  End If
  
  CSVStringToArray2D = CollectionToArray2D(oLinesCollection, oLinesColumnMax, Quote)
End Function

Private Function DetectLineEndingCode(InputStr As String) As String
  Dim LineEndingCode As Variant
  LineEndingCode = Array(vbCrLf, vbCr, vbLf)
             
  Dim No As Long
  For No = LBound(LineEndingCode) To UBound(LineEndingCode)
    If InStr(InputStr, LineEndingCode(No)) > 0 Then
      DetectLineEndingCode = LineEndingCode(No)
      Exit Function
    End If
  Next

  MsgBox "���s�R�[�h�̎������肪�ł��܂���ł���"
  End

End Function

Private Function CollectionToArray2D(oLinesCollection As Collection, oLinesColumnMax As Long, Quote As String) As Variant
  Dim oData As Variant
  ReDim oData(1 To oLinesCollection.Count, 1 To oLinesColumnMax) As Variant
  
  Dim oY As Long
  oY = 1
  
  Dim oLine As Variant
  For Each oLine In oLinesCollection
    Dim oX As Long
    oX = 1
    
    Dim oCell As Variant
    For Each oCell In oLine
      oData(oY, oX) = oCell
      oX = oX + 1
    Next
    
    oY = oY + 1
  Next

  CollectionToArray2D = oData
End Function


Private Function RemoveQuotes(ByVal Data As Variant, ByVal Quote As String) As Variant
  
'�擪�ƍŌオQuote�̏ꍇ
  If Left(Data, 1) = Quote And Right(Data, 1) = Quote Then
    
    '�擪�ƍŌ��Quote���J�b�g
    Data = Mid(Data, 2, Len(Data) - 2)
    
    '2�A��Quote��1��Quote�ɏC��
    Data = Replace(Data, Quote & Quote, Quote)
  End If
  
  RemoveQuotes = Data
End Function


'�t�@�C���ǂݍ���
'����I�����A�ǂݍ���String��Ԃ�
Public Function ReadFileToString(Filepath As String, Optional Charset As String = "_autodetect_all") As String
  If Dir(Filepath) = "" Then
    MsgBox "�t�@�C��:" & Filepath & "�����݂��܂���"
    End
  End If
   
'utf-8n�̂ݓ��ʑΉ�
  If Charset = "utf-8n" Then
    Charset = "utf-8"
  End If
   
  Dim ST As Object
  Set ST = CreateObject("ADODB.Stream")
  
  With ST
    .Mode = 3  'adModeReadWrite
    .Type = 2  'adTypeText
    .Charset = Charset
  
    .Open
    .LoadFromFile Filepath
    .Position = 0
    
    Dim buf As String
    buf = .ReadText(-1) 'adReadAll
  
    .Close
  End With
  
'utf-8�iBOM����j����������œǂݍ��񂾏ꍇ�̕␳�@�����ɂ���Ă͕s�v�ȉ\������
  If AscW(buf) = -257 Then    '-257 = &hFEFF
    buf = MidB(buf, 3, LenB(buf))
  End If
'utf-8�iBOM����j�␳�����܂�
    
  ReadFileToString = buf
End Function



'Lines�i�z��j�̒���iY�s�ڂ̃f�[�^�����o���B
'�������A���o�����s�ɁAQuote��������Ȃ��ꍇ�ɂ́A���̍s�����킹�ēǂݍ���
'��CSV�f�[�^�́A1�s�̒���Quote�͋�������͂��B
'�����ꍇ�ɂ́AQuote�̒��ɉ��s�����������Ă��āA���̍s�ƍ��킹��1�s�Ƃ��ĔF��������K�v������B
'������Quote������邩�ǂ����̔�����s��

Private Function GetLine(ByRef Lines() As String, ByRef iY As Long, ByRef LineEndingCode As String, ByRef Quote As String) As String
'����Function�ł́AiY�̒l�𑝂₵�Ă���
'ByRef�Ŏ󂯎���Ă���̂ŁA�Ăяo������iY���ω����邱�Ƃɒ��ӁI

  Dim Line As String
  Line = ""
  
  Dim LineEnding As String
  LineEnding = ""
    
  Dim QuoteCount As Long
  Do
    Line = Line & LineEnding & Lines(iY)
    QuoteCount = (QuoteCount + Len(Lines(iY)) - Len(Replace(Lines(iY), Quote, ""))) Mod 2
    
    iY = iY + 1

    If QuoteCount = 0 Or iY > UBound(Lines) Then
      Exit Do
    End If
    
    LineEnding = vbLf '�Z�������s��vbLf�̂��߁A���̉��s�R�[�h�ɂ�����炸vbLf���w��
  Loop
  
  If QuoteCount = 1 And iY > UBound(Lines) Then
    Err.Raise 65534, , "CSV�̃f�[�^�`�����s���ł��i���p�����K�؂ɕt����Ă��܂���j"
  End If

  GetLine = Line
  
End Function


Private Function ParseLine( _
         ByRef Line As String, _
         ByRef Regex As Object, _
         ByRef LineEndingCode As String, _
         ByRef Quote As String, _
         ByRef Delimiter As String, _
         ByRef SkipColumnDict As Object) As Collection
         
  Dim CSV As Collection
  Set CSV = New Collection
                 
  Dim Matches As Object
  Set Matches = Regex.Execute(Line & Delimiter)
  With Matches
    Dim Column As Long
    For Column = 1 To .Count
      If Not SkipColumnDict.Exists(CStr(Column)) Then
        CSV.Add RemoveQuotes(.Item(Column - 1).SubMatches(0), Quote)
      End If
    Next
  End With
  
  Set ParseLine = CSV
End Function


Private Function getRegexPattern(ByRef Delimiter As String, ByRef Quote As String) As Object
    
  Dim NotQuoted As String
  NotQuoted = "[^" & Delimiter & Quote & "]*"
' ���K�\����   [^,"]*
  
  
  Dim InsideOfQuote As String
  InsideOfQuote = "(?:[^" & Quote & "]|" & Quote & Quote & ")*"
' ���K�\����       (?:[^"]|"")*
   
  Dim Quoted As String
  Quoted = Quote & InsideOfQuote & Quote
' ���K�\����  "(?:[^"]|"")*"

  Dim Re As Object
  Set Re = CreateObject("VBScript.RegExp")
  Re.Global = True
  Re.MultiLine = False
  Re.Pattern = "(" & NotQuoted & "|" & Quoted & ")" & Delimiter
' ���K�\����   ([^,"]*|"(?:[^"]|"")*"),

  Set getRegexPattern = Re

End Function


Public Function NumberFormatParams(ParamArray Param()) As Object
  Dim Dict As Object
  Set Dict = CreateObject("Scripting.Dictionary")
  
  If UBound(Param) Mod 2 = 0 Then
    MsgBox "NumberFormatParams�̈����̐��͋����ɂ��Ă�������"
    End
  End If
  
  Dim No As Long
  For No = LBound(Param) To UBound(Param) Step 2
    Dict(Param(No)) = Param(No + 1)
  Next

  Set NumberFormatParams = Dict
End Function







