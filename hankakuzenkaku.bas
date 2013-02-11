Attribute VB_Name = "Module1"
Sub �S�p��()
    Dim c As Range
    
    For Each c In ActiveSheet.UsedRange  '�V�[�g�Ŏg�p�����Ă����Z�����ׂĂ��Ώۂ�
        If c.Value <> "" Then c.Value = ���p�J�i�����S�p���Ȃɕϊ�(c.Value)
    Next
End Sub

Function ���p�J�i�����S�p���Ȃɕϊ�(str As String)
    Dim buf As String
    Dim re As RegExp
    Dim m As Match
    
    Set re = New RegExp
    re.pattern = "[�-�]+" '���p�J�i�̃p�^�[��0xA1-0xDF(Wikipedia:���p�J�i���Q��)
    re.global = true

'    buf = StrConv(str, vbNarrow) '�p�������p�ɂ������ꍇ��
    buf = str
    For Each m In re.Execute(buf)
        buf = Replace(buf, m, StrConv(m, vbWide))
    Next
    ���p�J�i�����S�p���Ȃɕϊ� = buf
End Function
