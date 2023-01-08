Option Explicit

'' load external vbs.
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsConstants.vbs").ReadAll()

''Dim wdMainTextStory
const wdMainTextStory = 1 '' main story

Dim vbTab, vbCrLf, vbLf
vbTab = Chr(9)
vbCrLf = Chr(13) & Chr(10)
vbLf = Chr(10)

dim outputFileName
dim inputFileName
inputFileName = "C:\home2\JP\training.wd"
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


 
    Public WordApp ''as Word.Application
    Public WordDoc ''as Word.Document

    Sub CreateWord(lines)

        Set WordApp = CreateObject("Word.Application")  ' CreateObject関数でWordをセット
        With WordApp
            ''.Visible = True 'Wordを起動する（表示）
            '' Set WordDoc = .Documents.Add 'Wordを新規作成
            '' Set WordDoc = .Documents.Open("C:\home\見出しサンプル.docx")
            Set WordDoc = .Documents.Add("C:\home\sample-heder.docx")
            WordDoc.StoryRanges(wdMainTextStory).Delete
        End With

        Dim i ''as Long
        Dim params ''As String
        Dim message ''as String
        For i = 1 To lines.Count
            if i mod 10 = 0 Then
                log i, i/lines.Count*100, ""
            End if
            If lines.item(i) <> "" Then
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
                        Dim mergeInfo ''As String()
                        arrayInfo = getTableInfoArray(lines, i, columnWith, mergeInfo)
                        Call Me.AddTable(arrayInfo, columnWith, mergeInfo)
                    Case "image"
                        Call Me.AddImage(params(1))
                    Case "text"
                        Me.AddText (params(1))
                    Case "dddd"
                End Select
            End If
            
            ''DoEvents
        Next

        WordApp.Visible = True
        if Err.Description <> "" Then
            MsgBox Err.Description
        End If
    End Sub

    ''Function getTableInfoArray(lines As clsString, iCurrent As Long, columnWith() As String) As String()
    Function getTableInfoArray(lines, iCurrent, columnWith, mergeInfo) ''As String()
        Dim i ''as Long
        Dim iInfo ''as Long

        Dim strSplit ''As String
        strSplit = Split(lines.Item(iCurrent), vbTab)
        
        '' get create table and rows , columns
        Dim tableInfo ''As String()
        ReDim tableInfo(strSplit(1) - 1, strSplit(2) - 1)
        ReDim mergeInfo(strSplit(1) - 1, strSplit(2) - 1)
        ReDim columnWith(strSplit(2) - 1)
        Dim cellCount ''as Long
        cellCount = CLng(strSplit(1)) * CLng(strSplit(2))

        '' get table info
        Dim infoCount ''as Long
        infoCount = Split(lines.Item(iCurrent + 1), vbTab)(1)
        Dim strColumnWidth ''as String
        '' column info is separate by comma
        strColumnWidth = "1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"
        For i = 1 To infoCount
            strColumnWidth = split(lines.Item(iCurrent + 2), vbTab)(1)
        Next

        For i = 0 To UBound(columnWith)
            '' column info is separate by comma
            columnWith(i) = Split(strColumnWidth, ",")(i)
        Next

        '' create table
        For i = iCurrent + 1 + infoCount + 1 To iCurrent + cellCount + infoCount + 1
            strSplit = Split(lines.Item(i), vbTab)
            tableInfo(CLng(strSplit(1)), CLng(strSplit(2))) = strSplit(3)
        Next

        iCurrent = iCurrent + cellCount + 1
        getTableInfoArray = tableInfo

        '' tableMarge
        For i = iCurrent + 1  To 32000
            strSplit = Split(lines.Item(i), vbTab)
            if ubound(strSplit) > 4 then
                if (strSplit(5)) <> "" then
                    if strSplit(1) = strSplit(3) and strSplit(1) = strSplit(3) Then
                        '
                    else
                        mergeInfo(CLng(strSplit(1)), CLng(strSplit(2))) = strSplit(3) & "," &  strSplit(4)
                    end if
                End if
            else
                exit for
            End if
            iCurrent = iCurrent + 1
        Next
    End Function


    ''Sub AddTitle(ByRef mainTitle As String, ByRef subTitle As String)
    Sub AddTitle(ByRef mainTitle, ByRef subTitle)
        WordApp.Selection.TypeText mainTitle
        WordApp.Selection.Style = WordDoc.Styles("表題")
        WordApp.Selection.TypeText vbLf
        WordApp.Selection.TypeText subTitle
        WordApp.Selection.Style = WordDoc.Styles("副題")
        WordApp.Selection.TypeText vbLf
        WordApp.Selection.Style = WordDoc.Styles("body1")
    End Sub

    ''Sub AddHead(ByVal head As Long, ByRef text As String)
    Sub AddHead(ByVal head, ByRef text)
        'WordApp.Activate
        WordApp.Selection.TypeText text
        Dim strHead ''as String
        strHead = "見出し " & CStr(head)
        WordApp.Selection.Style = WordDoc.Styles(strHead)
        WordApp.Selection.TypeText vbLf
        WordApp.Selection.Style = WordDoc.Styles("body1")
    End Sub

    '' Sub AddOderList(ByVal head As Long, ByRef text As String)
    Sub AddOderList(ByVal head, ByRef text)
        WordApp.Selection.TypeText text
        Dim strHead ''as String
        strHead = "見出し " & CStr(head + 7)
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
        WordApp.Selection.InlineShapes.AddPicture imagePath
        WordApp.Selection.Style = WordDoc.Styles("body1")
        WordApp.Selection.TypeText vbLf
    End Sub

    ''Sub AddTable(ByRef table() As String, columnWith() As String, mergeInfo() as String)
    Sub AddTable(table(), columnWith(), mergeInfo())
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
            tableWidthSettings = tableWidthSettings + CSng(columnWith(y - 1))
        Next
        
        'set table size
        For y = 1 To tableColumns
            oTable.columns(y).Width = (tableWidth * 0.97) * CSng(columnWith(y - 1)) / tableWidthSettings



        Next
        
        oTable.Style = "styleN"
        dim MergeEnd
        dim cellRange
        
        For x = 1 To tableRows
            For y = 1 To tableColumns
                oTable.Cell(x, y).Range.text = table(x - 1, y - 1)

                ' if  mergeInfo(x-1, y-1) <> "" Then
                '     MergeEnd = split(mergeInfo(x-1, y-1), ",")
                '     WordDoc.Range(oTable.Cell(x, y).Range.Start, oTable.Cell(clng(MergeEnd(0)), clng(MergeEnd(1))).Range.End).Cells.Merge
                '     ''set cellRange = oTable.Cell(x, y)
                '     ''cellRange.Merge oTable.Cell(clng(MergeEnd(0)), clng(MergeEnd(1)))
                ' end if
            Next
        Next

        For x = 1 To tableRows
            For y = 1 To tableColumns
                ' oTable.Cell(x, y).Range.text = table(x - 1, y - 1)

                WScript.echo mergeInfo(x-1, y-1)
                if  mergeInfo(x-1, y-1) <> "" Then
                    MergeEnd = split(mergeInfo(x-1, y-1), ",")
                    'WordDoc.Range(oTable.Cell(1, 1), oTable.Cell(1, 2)).Cells.Merge
                    'oTable.Cell(x, y).Merge oTable.Cell(clng(MergeEnd(0)), clng(MergeEnd(1)))
                    WScript.echo "Merge", x-1, y-1, MergeEnd(0), MergeEnd(1)
                    oTable.Cell(x, y).Merge oTable.Cell(MergeEnd(0)+1, MergeEnd(1)+1)
                end if
            Next
        Next


        oTable.Style = "styleN"
        
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
                ' 既定値。現在の位置から EOS マーカー方向に、すべてのバイトをストリームから読み取ります。
                'これは、バイナリ ストリーム (Type は adTypeBinary) に唯一有効な StreamReadEnum 値です。
                ' adReadLine' -2
                ' ストリームから次の行を読み取ります (LineSeparator プロパティで指定)。
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
        'Microsoft ActiveX Data Objects 2.5 Libraryと
        ' WriteText str, 1 => add a newline
        ' WriteText str, 0 => add no newline
        Dim writeStream

        ' 文字コードを指定してファイルをオープン
        Set writeStream = CreateObject("ADODB.Stream")
        writeStream.Charset = "UTF-8"
        writeStream.Open

        '実際の中身の書き込み
        Dim i, items
        items = m_Files.items
        For i = 0 to m_Files.Count -1
            writeStream.WriteText m_Files.item(i)
            writeStream.WriteText "", 1
        next
        ' ファイルに書き込み
        writeStream.SaveToFile filename, 2 'adSaveCreateOverWrite:2

        ' ファイルをクローズ
        writeStream.Close
        Set writeStream = Nothing
    End Sub
End Class


Sub Log(m1, m2, m3)
    WScript.Echo "===>" & cstr(m1) & ":: " & cstr(m2) & ":: " & cstr(m3)
End Sub


