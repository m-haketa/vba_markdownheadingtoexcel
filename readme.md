# このファイルについて

## 目的

markdownファイルから見出し（#、##、###）を抜き出して、エクセルファイル上に目次を作成するマクロです。


## 使い方

エクセルファイルを開き、マクロを有効化したうえで、A1セルあたりにあるボタンを押してください。
markdownファイルを選択すると、自動で目次が作成されます。



# その他

ソースコード中の下記の行は、コメントのようにLen関数をStr関数に変更すると、**（私の手元の環境では）エクセルが強制終了します**。

```m_DataRange.Cells(Row, HeadingLevel).Value = Mid(Line, HeadingLevel + 2, Len(Line)) 'Len→Strに変えると落ちる```

原因がわかれば、教えてください。