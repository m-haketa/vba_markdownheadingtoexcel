Attribute VB_Name = "ReadCSV"
Option Explicit

'NumberFormats はDictionary（Key：列番号、Value：ユーザ定義書式）を指定する
'Charsetは初期設定の他、Shift_JIS、euc-jp、utf-8、utf-8nなどが設定できる

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


'NumberFormatsはDictionary前提
Private Function getColumnDataTypes(TextColumns As String, SkipColumns As String, ColumnNumberFormats As Object, MaxColumnNumber As Long) As Variant
  
  Dim SkipColumnDict As Object
  Set SkipColumnDict = ParseColumnsSelectString(SkipColumns)
  
'いったん、仮にSkipColumnDictがないものとして配列を作成
  Dim FormatData As Variant
  ReDim FormatData(1 To MaxColumnNumber + SkipColumnDict.Count)
  
  Dim C As Long
  For C = 1 To MaxColumnNumber + SkipColumnDict.Count
    FormatData(C) = "G/標準"
  Next
  
  Dim Column As Variant
  
  Dim TextColumnDict As Object
  Set TextColumnDict = ParseColumnsSelectString(TextColumns)
  
  For Each Column In TextColumnDict
    FormatData(CLng(Column)) = "@"
  Next


'より詳細な表示形式指定がある場合には、上書き
  If Not ColumnNumberFormats Is Nothing Then
    For Each Column In ColumnNumberFormats.Keys
      FormatData(Column) = ColumnNumberFormats(Column)
    Next
  End If
  
  
'Skipする列の補正を入れる
'Skip対象の列がある場合には、書式との対応がずれるので補正を入れる
'たとえば、CSVが4列で、書式が「文字列,文字列,skip,標準」という指定になっている場合。
'CSVの4列目の書式として指定している値（標準）を、シートの3列目に設定しないといけない
   
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


'LineEndingCodeには、vbCrLf vbCr vbLfなどを指定する
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


'CSVデータを格納した文字列を解析して2次元配列を返す
'LineEndingCodeには、vbCrLf vbCr vbLfなどを指定する
Public Function CSVStringToArray2D(ByRef InputStr As String, _
    Optional ByRef SkipColumns As String = "", _
    Optional ByVal ReadHeader As Boolean = True, _
    Optional ByRef Delimiter As String = ",", _
    Optional ByRef Quote As String = """", _
    Optional ByRef LineEndingCode As String = mc_AutoDetectString) As Variant
    
  If Len(Quote) > 1 Then
    MsgBox "引用符は1文字で指定してください"
    End
  End If
  
  If LineEndingCode = mc_AutoDetectString Then
    LineEndingCode = DetectLineEndingCode(InputStr)
  End If
  
'とりあえず、行数の最大値を取得するため、暫定的に入力データを改行コードごとに配列に分割
  Dim iLines() As String
  iLines = Split(InputStr, LineEndingCode)

  Dim oLinesCollection As Collection
  Set oLinesCollection = New Collection
   
  Dim Re As Object
  Set Re = getRegexPattern(Delimiter, Quote)
  
  Dim SkipColumnDict As Object
  Set SkipColumnDict = ParseColumnsSelectString(SkipColumns)
  
  
  Dim oLine As Variant  '1次元配列が入る
  Dim oLinesColumnMax As Long
  
  'iLinesの現在処理対象行
  Dim iY As Long
  iY = LBound(iLines)
  
  'CSVを解析した結果をいったんCollection（oLine）に格納
  '最大列数をoLinesColumnMaxに格納
  Do While iY <= UBound(iLines)
    
    'getLineの呼び出し後、iYが変化する（増える）ので注意！
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

  MsgBox "改行コードの自動判定ができませんでした"
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
  
'先頭と最後がQuoteの場合
  If Left(Data, 1) = Quote And Right(Data, 1) = Quote Then
    
    '先頭と最後のQuoteをカット
    Data = Mid(Data, 2, Len(Data) - 2)
    
    '2連続Quoteを1つのQuoteに修正
    Data = Replace(Data, Quote & Quote, Quote)
  End If
  
  RemoveQuotes = Data
End Function


'ファイル読み込み
'正常終了時、読み込んだStringを返す
Public Function ReadFileToString(Filepath As String, Optional Charset As String = "_autodetect_all") As String
  If Dir(Filepath) = "" Then
    MsgBox "ファイル:" & Filepath & "が存在しません"
    End
  End If
   
'utf-8nのみ特別対応
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
  
'utf-8（BOMあり）を自動判定で読み込んだ場合の補正　※環境によっては不要な可能性あり
  If AscW(buf) = -257 Then    '-257 = &hFEFF
    buf = MidB(buf, 3, LenB(buf))
  End If
'utf-8（BOMあり）補正ここまで
    
  ReadFileToString = buf
End Function



'Lines（配列）の中のiY行目のデータを取り出す。
'ただし、取り出した行に、Quoteが奇数個しかない場合には、次の行も合わせて読み込む
'※CSVデータは、1行の中にQuoteは偶数個あるはず。
'奇数個ある場合には、Quoteの中に改行文字が入っていて、次の行と合わせて1行として認識させる必要がある。
'そこでQuoteが奇数個あるかどうかの判定を行う

Private Function GetLine(ByRef Lines() As String, ByRef iY As Long, ByRef LineEndingCode As String, ByRef Quote As String) As String
'このFunctionでは、iYの値を増やしている
'ByRefで受け取っているので、呼び出し元でiYが変化することに注意！

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
    
    LineEnding = vbLf 'セル内改行はvbLfのため、元の改行コードにかかわらずvbLfを指定
  Loop
  
  If QuoteCount = 1 And iY > UBound(Lines) Then
    Err.Raise 65534, , "CSVのデータ形式が不正です（引用符が適切に付されていません）"
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
' 正規表現例   [^,"]*
  
  
  Dim InsideOfQuote As String
  InsideOfQuote = "(?:[^" & Quote & "]|" & Quote & Quote & ")*"
' 正規表現例       (?:[^"]|"")*
   
  Dim Quoted As String
  Quoted = Quote & InsideOfQuote & Quote
' 正規表現例  "(?:[^"]|"")*"

  Dim Re As Object
  Set Re = CreateObject("VBScript.RegExp")
  Re.Global = True
  Re.MultiLine = False
  Re.Pattern = "(" & NotQuoted & "|" & Quoted & ")" & Delimiter
' 正規表現例   ([^,"]*|"(?:[^"]|"")*"),

  Set getRegexPattern = Re

End Function


Public Function NumberFormatParams(ParamArray Param()) As Object
  Dim Dict As Object
  Set Dict = CreateObject("Scripting.Dictionary")
  
  If UBound(Param) Mod 2 = 0 Then
    MsgBox "NumberFormatParamsの引数の数は偶数個にしてください"
    End
  End If
  
  Dim No As Long
  For No = LBound(Param) To UBound(Param) Step 2
    Dict(Param(No)) = Param(No + 1)
  Next

  Set NumberFormatParams = Dict
End Function







