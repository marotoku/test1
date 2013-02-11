Attribute VB_Name = "Module1"
Sub 全角化()
    Dim c As Range
    
    For Each c In ActiveSheet.UsedRange  'シートで使用されているセルすべてを対象に
        If c.Value <> "" Then c.Value = 半角カナだけ全角かなに変換(c.Value)
    Next
End Sub

Function 半角カナだけ全角かなに変換(str As String)
    Dim buf As String
    Dim re As RegExp
    Dim m As Match
    
    Set re = New RegExp
    re.pattern = "[｡-ﾟ]+" '半角カナのパターン0xA1-0xDF(Wikipedia:半角カナを参照)
    re.global = true

'    buf = StrConv(str, vbNarrow) '英数も半角にしたい場合に
    buf = str
    For Each m In re.Execute(buf)
        buf = Replace(buf, m, StrConv(m, vbWide))
    Next
    半角カナだけ全角かなに変換 = buf
End Function
