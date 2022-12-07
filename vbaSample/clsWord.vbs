Option Explicit

Dim wdMainTextStory
wdMainTextStory = 1

Dim vbTab, vbCrLf, vbLf
vbTab = Chr(9)
vbCrLf = Chr(13) & Chr(10)
vbLf = Chr(10)

dim outputFileName
dim inputFileName
inputFileName = "C:\Users\maru\Desktop\word-tips\JP\md2docx.wd"
''outputFileName = "C:\projects\markdown-to-docx\createDocsEx.vbs"
dim fileList
set fileList = new XFiles

'fileList.Add 1, "title" & vbTab & "title" & vbTab & "subtitle"
call fileList.Load(inputFileName)
''call fileList.WriteToFile(outputFileName)

dim w
set w = new XWord
call w.CreateWord(fileList)

Class XWord

' wdCommentsStory	4	
' �R�����g �X�g�[���[

' wdEndnoteContinuationNoticeStory	17	
' �����r���p�����̒��̃X�g�[���[

' wdEndnoteContinuationSeparatorStory	16	
' �����r���p�����̋��E���̃X�g�[���[

' wdEndnoteSeparatorStory	15	
' �����r���̋��E���̃X�g�[���[

' wdEndnotesStory	3	
' �����r���̃X�g�[���[

' wdEvenPagesFooterStory	8	
' �����y�[�W�̃t�b�^�[�̃X�g�[���[

' wdEvenPagesHeaderStory	6	
' �����y�[�W�̃w�b�_�[�̃X�g�[���[

' wdFirstPageFooterStory	11	
' �擪�y�[�W�̃t�b�^�[�̃X�g�[���[

' wdFirstPageHeaderStory	10	
' �擪�y�[�W�̃w�b�_�[�̃X�g�[���[

' wdFootnoteContinuationNoticeStory	14	
' �r���p�����̒��̃X�g�[���[

' wdFootnoteContinuationSeparatorStory	13	
' �r���p�����̋��E���̃X�g�[���[

' wdFootnoteSeparatorStory	12	
' �r���̋��E���̃X�g�[���[

' wdFootnotesStory	2	
' �r���̃X�g�[���[
' Dim wdMainTextStory
' wdMainTextStory = 1	
' ���C�� �e�L�X�g �X�g�[���[

' wdPrimaryFooterStory	9	
' ��t�b�^�[�̃X�g�[���[

' wdPrimaryHeaderStory	7	
' ��w�b�_�[�̃X�g�[���[

' wdTextFrameStory	5	
' ���C�A�E�g�g�̃X�g�[���[
 
    Public WordApp ''as Word.Application
    Public WordDoc ''as Word.Document

    Sub CreateWord(lines)

        Set WordApp = CreateObject("Word.Application")  ' CreateObject�֐���Word���Z�b�g
        With WordApp
            ''.Visible = True 'Word���N������i�\���j
            '' Set WordDoc = .Documents.Add 'Word��V�K�쐬
            '' Set WordDoc = .Documents.Open("C:\home\���o���T���v��.docx")
            Set WordDoc = .Documents.Add("C:\home\sample-heder.docx")
            WordDoc.StoryRanges(wdMainTextStory).Delete
        End With

        Dim i ''as Long
        Dim params ''As String
        Dim message ''as String
        For i = 1 To lines.Count
            If lines.item(i) <> "" Then
                WScript.Echo lines.item(i)
                message = lines.item(i)
                params = Split(message, vbTab)
                Select Case params(0)
                    Case "title"
                        dim mainTitle
                        dim subTitle
                        Call Me.AddTitle(params(1), params(2))
                    Case "section"
                        Call Me.AddHead(params(1), params(2))
                    Case "OderList"
                        Call Me.AddOderList(params(1), params(2))
                    Case "NormalList"
                        Call Me.AddNormalList(params(1), params(2))
                    Case "tableCreate"
                        Dim columnWith ''As String()
                        Dim arrayInfo ''As String()
                        arrayInfo = getTableInfoArray(lines, i, columnWith)
                        Call Me.AddTable(arrayInfo, columnWith)
                    Case "image"
                        Call Me.AddImage(message)
                    Case "text"
                        Me.AddText (message)
                    Case "dddd"
                    
                End Select
            End If
            
            ''DoEvents
        Next

        WordApp.Visible = True
        MsgBox Err.Description
    End Sub



    
    ''Function getTableInfoArray(lines As clsString, iCurrent As Long, columnWith() As String) As String()
    Function getTableInfoArray(lines, iCurrent, columnWith) ''As String()
        Dim i ''as Long
        Dim iInfo ''as Long

        Dim strSplit ''As String
        strSplit = Split(lines.Item(iCurrent), vbTab)
        
        '' get create table and rows , columns
        Dim tableInfo ''As String()
        ReDim tableInfo(strSplit(1) - 1, strSplit(2) - 1)
        ReDim columnWith(strSplit(2) - 1)
        Dim cellCount ''as Long
        cellCount = CLng(strSplit(1)) * CLng(strSplit(2))

        '' get table info
        Dim infoCount ''as Long
        infoCount = Split(lines.Item(iCurrent + 1), vbTab)(1)
        Dim strColumnWidth ''as String
        strColumnWidth = strColumnWidth & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab
        For i = 1 To infoCount
            strColumnWidth = RegGetMessage(lines.Item(iCurrent + 2))
        Next

        For i = 0 To UBound(columnWith)
            columnWith(i) = Split(strColumnWidth, vbTab)(i)
        Next

        '' create table
        For i = iCurrent + 1 + infoCount + 1 To iCurrent + cellCount + infoCount + 1
            strSplit = Split(lines.Item(i), vbTab)
            tableInfo(CLng(strSplit(1)), CLng(strSplit(2))) = strSplit(3)
        Next
        
        iCurrent = iCurrent + cellCount + 1
        getTableInfoArray = tableInfo
    End Function


    ''Sub AddTitle(ByRef mainTitle As String, ByRef subTitle As String)
    Sub AddTitle(ByRef mainTitle, ByRef subTitle)
        WordApp.Selection.TypeText mainTitle
        WordApp.Selection.Style = WordDoc.Styles("�\��")
        WordApp.Selection.TypeText vbLf
        WordApp.Selection.TypeText subTitle
        WordApp.Selection.Style = WordDoc.Styles("����")
        WordApp.Selection.TypeText vbLf
        WordApp.Selection.Style = WordDoc.Styles("body1")
    End Sub

    ''Sub AddHead(ByVal head As Long, ByRef text As String)
    Sub AddHead(ByVal head, ByRef text)
        'WordApp.Activate
        WordApp.Selection.TypeText text
        Dim strHead ''as String
        strHead = "���o�� " & CStr(head)
        WordApp.Selection.Style = WordDoc.Styles(strHead)
        WordApp.Selection.TypeText vbLf
        WordApp.Selection.Style = WordDoc.Styles("body1")
    End Sub

    '' Sub AddOderList(ByVal head As Long, ByRef text As String)
    Sub AddOderList(ByVal head, ByRef text)
        WordApp.Selection.TypeText text
        Dim strHead ''as String
        strHead = "���o�� " & CStr(head + 7)
        WordApp.Selection.Style = WordDoc.Styles(strHead)
        WordApp.Selection.TypeText vbLf
        WordApp.Selection.Style = WordDoc.Styles("body1")
    End Sub

    ''Sub AddNormalList(ByVal list As Long, ByRef text As String)
    Sub AddNormalList(ByVal list, ByRef text)
        WordApp.Selection.TypeText text
        Dim strHead ''as String
        strHead = "nlist" & CStr(list)
        WordApp.Selection.Style = WordDoc.Styles(strHead)
        WordApp.Selection.TypeText vbLf
        WordApp.Selection.Style = WordDoc.Styles("body1")
    End Sub

    ''Sub AddText(ByRef text As String)
    Sub AddText(ByRef text)
        WordApp.Selection.TypeText text
        WordApp.Selection.Style = WordDoc.Styles("body1")
        WordApp.Selection.TypeText vbLf
    End Sub

    ''Sub AddImage(ByRef imagePath As String)
    Sub AddImage(ByRef imagePath)
        ''WordApp.Selection.InlineShapes.AddPicture fileName:= imagePath, LinkToFile:=False, SaveWithDocument:= True
        WordApp.Selection.InlineShapes.AddPicture imagePath, False, False
        WordApp.Selection.Style = WordDoc.Styles("body1")
        WordApp.Selection.TypeText vbLf
    End Sub

    sub test()
        call WriteToFile()
        AddTable
    End sub

    ''Sub AddTable(ByRef table() As String, columnWith() As String)
    Sub AddTable(ByRef table(), columnWith())
        WordApp.ActiveDocument.Range.InsertParagraphAfter
        'create table and asign it to variable
        Dim oTable ''as Word.table
        Dim x ''as Long
        Dim y ''as Long
        Dim tableRows ''as Long
        Dim tableColumns ''as Long
        tableRows = UBound(table, 1) + 1
        tableColumns = UBound(table, 2) + 1
        'Set oTable = WordApp.ActiveDocument.tables.Add(Range:=WordApp.ActiveDocument.Paragraphs.Last.Range, NumRows:=tableRows, NumColumns:=tableColumns, DefaultTableBehavior:=1)
        Set oTable = WordApp.ActiveDocument.tables.Add(WordApp.ActiveDocument.Paragraphs.Last.Range, tableRows, tableColumns, 1)
            
        oTable.Style = "styleN"
        
        Dim tableWidth ''as Single
        Dim tableWidthSettings ''as Single
        tableWidth = 0
        tableWidthSettings = 0
        For y = 1 To tableColumns
            tableWidth = tableWidth + oTable.columns(y).Width
            tableWidthSettings = tableWidthSettings + CDbl(columnWith(y - 1))
        Next
        
        'set table size
        For y = 1 To tableColumns
            oTable.columns(y).Width = (tableWidth * 0.97) * CDbl(columnWith(y - 1)) / tableWidthSettings
        Next
        
        oTable.Style = "styleN"
        
        For x = 1 To tableRows
            For y = 1 To tableColumns
                oTable.Cell(x, y).Range.text = table(x - 1, y - 1)
            Next
        Next
        
        Dim r
        Set r = oTable.Range
        r.SetRange oTable.Range.End + 1, oTable.Range.End + 1
        r.Select
    End Sub
End Class

Class XFiles
    Dim m_Files
    Dim FSO

    Private Sub Class_Initialize
        Set m_Files = CreateObject("Scripting.Dictionary")
        Set FSO = CreateObject("Scripting.FileSystemObject")
    End Sub

    Public Function Files()
        set Files = m_Files
    End Function

    Public Sub Clear()
        Set m_Files = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub Add(key, value)
        m_Files.Add key, value
    End Sub

    Public Function Count()
        Count = m_Files.Count
    End Function

    Public Function Item(i)
        Item = m_Files.Item(i)
    End Function

    Public Sub Remove(key)
        m_Files.remove key
    End Sub

    Public Sub Load(filePath)
        Dim pathFile
        dim ppath
        Dim currentLine, tmpSplitLine, comment
        Dim firstChar
        Dim currentLineNo
        Set m_Files = CreateObject("Scripting.Dictionary")

        If FSO.FileExists(filePath) = false Then
            exit Sub
        End If
        With CreateObject("ADODB.Stream")
            .Charset = "UTF-8"
            .Open
            .LoadFromFile filePath
            Do Until .EOS
                ' adReadAll ' -1
                ' ����l�B���݂̈ʒu���� EOS �}�[�J�[�����ɁA���ׂẴo�C�g���X�g���[������ǂݎ��܂��B
                '����́A�o�C�i�� �X�g���[�� (Type �� adTypeBinary) �ɗB��L���� StreamReadEnum �l�ł��B
                ' adReadLine' -2
                ' �X�g���[�����玟�̍s��ǂݎ��܂� (LineSeparator �v���p�e�B�Ŏw��)�B
                '[; ] relative path | comment | info
                currentLine = .ReadText(-2)
                currentLineNo = currentLineNo + 1
                m_Files.add currentLineNo, currentLine
            Loop
            .Close
        End with
    End Sub

    Sub WriteToFile(filename)
        'Dim writeStream As ADODB.Stream
        'Microsoft ActiveX Data Objects 2.5 Library��
        ' WriteText str, 1 => add a newline
        ' WriteText str, 0 => add no newline
        Dim writeStream

        ' �����R�[�h���w�肵�ăt�@�C�����I�[�v��
        Set writeStream = CreateObject("ADODB.Stream")
        writeStream.Charset = "UTF-8"
        writeStream.Open

        '���ۂ̒��g�̏�������
        Dim i, items
        items = m_Files.items
        For i = 0 to m_Files.Count -1
            writeStream.WriteText m_Files.item(i)
            writeStream.WriteText "", 1
        next
        ' �t�@�C���ɏ�������
        writeStream.SaveToFile filename, 2 'adSaveCreateOverWrite:2

        ' �t�@�C�����N���[�Y
        writeStream.Close
        Set writeStream = Nothing
    End Sub
End Class



