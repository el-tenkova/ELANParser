Attribute VB_Name = "ELANconv"

Public Function startsWithDigits() As Object
    ' Create a regular expression object.
    Set objRegExp = New RegExp

    'Set the pattern by using the Pattern property.
    objRegExp.Pattern = "^(\d+)\.(\d+)\s*([А-Я]*)\s*.\s*([А-Я].+)$"

    ' Set Case Insensitivity.
    objRegExp.IgnoreCase = True

    'Set global applicability.
    objRegExp.Global = True

    'Test whether the String can be compared.
     Set startsWithDigits = objRegExp
    
End Function

Public Function startsWithCaps() As Object
    ' Create a regular expression object.
    Set objRegExp = New RegExp

    'Set the pattern by using the Pattern property.
    objRegExp.Pattern = "^([А-Я]*)\s*.\s*([А-Я].+$)"

    ' Set Case Insensitivity.
    objRegExp.IgnoreCase = True

    'Set global applicability.
    objRegExp.Global = True

    'Test whether the String can be compared.
     Set startsWithCaps = objRegExp
    
End Function
Public Function createObjToNormalize() As Object
    ' Create a regular expression object.
    Set objRegExp = New RegExp

    'Set the pattern by using the Pattern property.
    objRegExp.Pattern = "[aceiopxy" & ChrW$(&HF6) & ChrW$(&HFF) & ChrW$(&H4B7) & "]"

    ' Set Case Insensitivity.
    objRegExp.IgnoreCase = True

    'Set global applicability.
    objRegExp.Global = True

    'Test whether the String can be compared.
     Set createObjToNormalize = objRegExp
    
End Function
Public Function createReplToNormalize() As Collection
    Dim repl As New Collection
    repl.Add "а", "a"
    repl.Add "с", "c"
    repl.Add "е", "e"
    repl.Add ChrW$(&H456), "i"
    repl.Add "о", "o"
    repl.Add "р", "p"
    repl.Add "х", "x"
    repl.Add "у", "y"
    repl.Add ChrW$(&H4E7), ChrW$(&HF6)
    repl.Add ChrW$(&H4F1), ChrW$(&HFF)
    repl.Add ChrW$(&H4CC), ChrW$(&H4B7)
'   'a': 'а', 'c': 'с', 'e': 'е', 'i': '\u0456', 'o': 'о', 'p': 'р', 'x': 'х', 'y': 'у', '\u00F6': '\u04E7', '\u00FF': '\u04F1', '\u04B7': '\u04CC'
    Set createReplToNormalize = repl
    
End Function

Sub SaveAsCSV()
Dim p As Long
Dim i As Long
Dim j As Long
Dim c As Long
Dim prevTime As Long
Dim nextTime As Long
Dim lastMin As String

Dim para As Paragraph

Dim Title As String

Dim fsT As Object
Set fsT = CreateObject("ADODB.Stream")
fsT.Type = 2
fsT.Charset = "utf-8"
fsT.Open

Dim esT As Object
Set esT = CreateObject("ADODB.Stream")
esT.Type = 2
esT.Charset = "utf-8"
esT.Open

Dim regexpObjDigits As Object
Dim regexpObjCaps As Object

p = ActiveDocument.Paragraphs.count
Set regexpObjDigits = startsWithDigits()
Set regexpObjCaps = startsWithCaps()

translation = ""
time_begin = ""
time_prev = ""
participant = ""
text = ""
Line = ""

timePrev = 0
timeNext = 0

c = 0
Set para = ActiveDocument.Paragraphs.item(1)
For i = 1 To p
    c = para.Range.words.count
    Title = ""
    wrd = ""
    For j = 1 To c
        Title = Title + para.Range.words.item(j)
    Next j

    If Len(Title) Then
        If regexpObjDigits.Test(Title) = True Then
            Debug.Print Title
            'fsT.WriteText Title
            'Get the matches.
            Set colMatches = regexpObjDigits.Execute(Title) ' Execute search.
            If time_begin <> "" Then
                Line = time_begin & ";"
                time_prev = time_begin
            End If
            For Each objMatch In colMatches   ' Iterate Matches collection.
                Index = 0
                For Each item In objMatch.SubMatches
                    If Index = 0 Then
                        time_begin = "00:" & item & ":"
                        timeNext = CInt(item) * 60
                        lastMin = item
                    Else
                        If Index = 1 Then
                            If time_prev = time_begin & item & ".00" Then
                                time_begin = time_begin & item & ".10"
                            Else
                                time_begin = time_begin & item & ".00"
                            End If
                            timeNext = timeNext + item
                            time_prev = time_begin
                            If Line <> "" Then
                                Line = Line & time_begin & ";" & text & ";" & translation
                                Line = Line & vbCrLf
                                
                                If timePrev > timeNext Then
                                    esT.WriteText "Ошибка во временном интервале: "
                                    esT.WriteText Line
                                End If
                                
                                fsT.WriteText Line
                                Line = ""
                                translation = ""
                                text = ""
                            End If
                            timePrev = timeNext
                        Else
                            If Index = 2 Then
                                participant = item
                            Else
                                If Index = 3 Then
                                    text = Mid$(item, 1, Len(item) - 1)
                                End If
                            End If
                        End If
                    End If
                    Index = Index + 1
                    Debug.Print item
                Next
            Next
           ' fsT.WriteText time_begin & vbCrLf
           ' fsT.WriteText participant & vbCrLf
           ' fsT.WriteText Text
        Else
            If regexpObjCaps.Test(Title) = True Then
                Debug.Print Title
                'fsT.WriteText Title
                'Get the matches.
                Set colMatches = regexpObjCaps.Execute(Title) ' Execute search.
                For Each objMatch In colMatches   ' Iterate Matches collection.
                    Index = 0
                    For Each item In objMatch.SubMatches
                        If Index = 1 Then
                            translation = item
                            'fsT.WriteText second
                        End If
                        Index = Index + 1
                        Debug.Print item
                    Next
                Next
            '    fsT.WriteText participant & vbCrLf
             '   fsT.WriteText Text
            
            End If
        End If
    End If
    Set para = para.Next()
Next i

If text <> "" Then
    t = CInt(lastMin) + 3
    time_next = Replace(time_begin, ":" & lastMin, ":" & Format$(t), 1, 1)
    Line = Line & time_begin & ";" & time_next & ";" & text & ";" & translation & vbCrLf
    fsT.WriteText Line
End If
NameCSV = ActiveDocument.Path & "\" & Mid$(ActiveDocument.name, 1, InStrRev(ActiveDocument.name, ".")) & "csv"
fsT.SaveToFile NameCSV, 2
fsT.Close
esT.SaveToFile ActiveDocument.Path & "\!!err.txt", 2
esT.Close

End Sub



Sub SaveAsEAF()
Dim t As Long
Dim i As Long
Dim j As Long
Dim c As Long

Dim tabl As table

Dim Title As String
Dim begin As LongLong

Dim fsT As Object
Set fsT = CreateObject("ADODB.Stream")
fsT.Type = 2
fsT.Charset = "utf-8"
fsT.Open

Dim esT As Object
Set esT = CreateObject("ADODB.Stream")
esT.Type = 2
esT.Charset = "utf-8"
esT.Open

Dim elProc As New ELANProc

t = 6 'ActiveDocument.Tables.count
' Запись заголовка файла
Call elProc.writeHeader(fsT)
' Запись временных интервалов
fsT.WriteText ("<TIME_ORDER>" & vbCrLf)
cells_cnt = 0
begin = 0
Dim time_intervals() As Integer
ReDim time_intervals(1 To t + 1)
time_intervals(1) = 1
j = 1
For i = 4 To t
    Set tabl = ActiveDocument.Tables.item(i)
    If tabl.Rows.count = 4 Then
        time_intervals(j + 1) = cells_cnt + tabl.Range.Cells.count + 1
        cells_cnt = cells_cnt + tabl.Range.Cells.count
        Call elProc.writeTimeSlotEx(time_intervals(j), time_intervals(j + 1), begin, fsT)
        j = j + 1
    End If
Next i
Call elProc.writeTimeSlotEx(time_intervals(j), time_intervals(j) + 4, begin, fsT)
fsT.WriteText ("</TIME_ORDER>" & vbCrLf)

Dim row As Object
Dim annoID As Long
annoID = 1
Dim delta As Integer
' Запись слоёв
Dim time1 As LongLong
Dim time2 As LongLong
Dim text As String
For Num = 1 To 4
    fsT.WriteText ("<TIER DEFAULT_LOCALE=""ru"" LINGUISTIC_TYPE_REF=""default"" TIER_ID=""Tier-")
    fsT.WriteText (Format$(Num - 1) & """")
    fsT.WriteText (">" & vbCrLf)
    delta = Num - 1
    timeJ = 1
    For i = 4 To t
        Set tabl = ActiveDocument.Tables.item(i)
        If tabl.Rows.count = 4 Then
            timeIdx = time_intervals(timeJ) + delta
            Set row = tabl.Rows.item(Num)
            For j = 1 To row.Cells.count
                text = row.Cells.item(j)
                time1 = timeIdx
                time2 = timeIdx + 3
                
                If Num = 4 Then
                    time2 = time_intervals(timeJ + 1)
                Else
                    If j = 1 Then
                        time2 = time2 + 1
                    End If
                End If
                Call elProc.writeAnno(Mid$(text, 1, Len(text) - 2), annoID, time1, time2, fsT)
                annoID = annoID + 1
                timeIdx = time2
            Next j
            timeJ = timeJ + 1
        Else
            esT.WriteText ("В таблице номер " & Format$(i) & " меньше четырёх строк" & vbCrLf)
            'esT.WriteText (Format$(timeIdx))
        End If
    Next i
    fsT.WriteText ("</TIER>" & vbCrLf)
Next Num
' Запись последнего блока файла
Call elProc.writeTail(fsT)
NameCSV = ActiveDocument.Path & "\" & Mid$(ActiveDocument.name, 1, InStrRev(ActiveDocument.name, ".")) & "eaf"
fsT.SaveToFile NameCSV, 2
fsT.Close
esT.SaveToFile ActiveDocument.Path & "\!!err.txt", 2
esT.Close

End Sub
Public Function Normalize(inWord As String, regExpObj As Object, repl As Collection) As String
    Dim str As String
    
    str = LCase$(inWord)
    Set colMatches = regExpObj.Execute(str) ' Execute search.
    For Each objMatch In colMatches   ' Iterate Matches collection.
            str = Replace$(str, objMatch, repl.item(objMatch))
    Next objMatch
    Normalize = str

End Function
Public Function wordsCount(para As Paragraph) As Long
    Dim punct As String
    Dim col As Long
    punct = ".,:;!?"
    col = 0
    c = para.Range.words.count - 1
    For i = 1 To c
        If InStr(punct, Mid(para.Range.words.item(i), 1, 1)) = 0 Then
            col = col + 1
        End If
    Next i
    wordsCount = col
End Function
Public Function ParseResponse(responsetext As String) As Collection
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim str As String
    Dim result As New Collection
    Dim idx As Long
    idx = 1
    str = responsetext
    i = InStr(str, "FOUND STEM:")
    Dim homonym As Collection
    Do While i <> 0
        Set homonym = New Collection
        str = Mid(str, i + 12)
        j = InStr(str, " ")
        i = InStr(str, Chr(10))
        'Debug.Print str
        
        homonym.Add Mid(str, 1, j), "form"
        homonym.Add Mid(str, j + 1, i - (j + 1)), "affixes"
        str = Mid(str, i + 1)
        homonym.Add Mid(str, 1, 1), "p_o_s"
        i = InStr(str, Chr(10))
    
        dict = Mid(str, 2, i - 2)
        j = InStr(dict, " " & ChrW(&H201B))
        homonym.Add Mid(dict, 1, j), "headword"
        k = InStr(dict, ChrW(&H2019))
        homonym.Add Mid(dict, j + 1, k + 1 - (j + 1)), "meaning"
        
        homonym.Add "0", "has-stem"
        If (Len(dict) > k + 1) Then
            homonym.Add Mid(dict, k + 2), "stem"
            homonym.Remove "has-stem"
            homonym.Add "1", "has-stem"
            
        End If
        
        result.Add homonym, CStr(idx)
        idx = idx + 1
        
        i = InStr(str, "FOUND STEM:")
    Loop
    Set ParseResponse = result
    
End Function

Sub DoParseKhakas()

    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open

    Dim punct As String
    punct = ".,:;!?"

    Dim strServ As String
    Dim strData As String
    Dim strResponse As String
    
    Dim newDoc As Document
    Dim theDoc As Document
    
    Dim para As Paragraph
    Dim Table1 As table
    
    Set theDoc = ActiveDocument
    Set newDoc = Documents.Add()
    Set Range1 = newDoc.Range(start:=0, End:=0)
   
    Dim regExpObj As Object
    Set regExpObj = createObjToNormalize()
    Dim repl As Object
    Set repl = createReplToNormalize()
    
    strServ = "http://khakas.altaica.ru/cgi-bin/suddenly.fs?Getparam=getval"
    strData = "&parse="

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    p = theDoc.Paragraphs.count
    c = 0
    Set para = theDoc.Paragraphs.item(1)
    For i = 1 To p
        c = para.Range.words.count

        If para.Range.Font.Italic = True Then
            If Table1.Rows.count < 4 Then
                Set row = Table1.Rows.Add()
                row.Cells.Merge
                row.Cells(1).Range.text = Mid(para.Range.text, 1, Len(para.Range.text) - 1)
            Else
                Selection.EndKey wdStory
                Selection.InsertAfter (para)
                Selection.EndKey wdStory
            End If
        Else
            If c > 1 Then
                Set Table1 = newDoc.Tables.Add(Range1, 2, c - 1)
                Table1.Borders.OutsideLineStyle = wdLineStyleSingle
                Table1.Borders.InsideLineStyle = wdLineStyleSingle
            End If
            Title = ""
            wrd = ""
            col = 1
            For j = 1 To c - 1
                If InStr(punct, Mid(para.Range.words.item(j), 1, 1)) = 0 Then
                    Table1.cell(1, col).Range.text = para.Range.words.item(j)
                    normWord = Normalize(para.Range.words.item(j), regExpObj, repl)
                    objHTTP.Open "post", strServ & strData & normWord, False
                    objHTTP.Send strData
                    If objHTTP.Status = 200 Then
                        Dim Response As Collection
                        Set Response = ParseResponse(objHTTP.responsetext)
                        'Table1.Cell(2, col).Range.text = strResponse
                    'fsT.WriteText (strResponse & vbCrLf)
                'Debug.Print strResponse
                    Else
                'Debug.Print "Ошибка"
                    End If
                    Dim homonym As Collection
                    
                    str1 = ""
                    'str2 = ""
                    For Each homonym In Response   ' Iterate Matches collection.
                        str1 = str1 & homonym.item("p_o_s") & homonym.item("headword") & homonym.item("meaning")
                        If homonym.item("has-stem") = "1" Then
                            str1 = str1 & " " & homonym.item("stem")
                        'Else
                        '    str1 = str1 & homonym.item("headword")
                        End If
                        str1 = str1 & vbCrLf
                        str1 = str1 & homonym.item("affixes") & vbCrLf
                        str1 = str1 & homonym.item("form") & vbCrLf
                        'str2 = str2 & homonym.item("meaning") & homonym.item("form") & vbCrLf
                        Table1.cell(2, col).Range.text = str1
                        'Table1.Cell(3, col).Range.text = str2
                        
                    Next homonym
                    col = col + 1
                    
                Else
                    
                    Table1.cell(1, col).Range.text = para.Range.words.item(j)
                    Table1.cell(1, col - 1).Merge Table1.cell(1, col)
                    Table1.cell(1, col - 1).Range.Find.Execute findtext:="^0013", replacewith:="", Replace:=wdReplaceAll

                    If Len(Table1.cell(2, col - 1).Range.text) > 2 Then
                        Table1.cell(2, col).Range.text = para.Range.words.item(j)
                        Table1.cell(2, col - 1).Merge Table1.cell(2, col)
                        Table1.cell(2, col - 1).Range.Find.Execute findtext:="^0013", replacewith:="", Replace:=wdReplaceAll
                    Else
                        Table1.cell(2, col - 1).Merge Table1.cell(2, col)
                    End If
                    'Table1.Cell(3, col - 1).Merge Table1.Cell(3, col)
                    
                End If
            Next j
            
            Selection.EndKey wdStory
            Selection.InsertAfter ("" & vbCrLf)
            Selection.EndKey wdStory
            Set Range1 = Selection.Range
            
        End If
        Set para = para.Next()
    Next i
    fsT.SaveToFile "c:\ELAN\test-parse.txt", 2
    fsT.Close
    
End Sub

