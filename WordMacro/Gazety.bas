Attribute VB_Name = "Gazety"
Function getSentCount(theDoc As Document) As Long
    Dim count As Long
    Dim cs As Long
    Dim cp As Long
    Dim para As Paragraph
    
    Set para = theDoc.Paragraphs.item(1)
    cp = theDoc.Paragraphs.count
    count = 0
    cs = 0
    For i = 1 To cp
        cs = 0
        If para.Range.Characters.count > 1 Then
            cs = para.Range.Sentences.count
        End If
        If cs <> 0 Then
            count = count + cs
        End If
        Set para = para.Next()
    Next i
    getSentCount = count
End Function
Sub writeColumn(theDoc As Document, number As Long, Table1 As table)
    Dim cs As Long
    Dim cp As Long
    Dim count As Long
        
    Dim para As Paragraph

    Set para = theDoc.Paragraphs.item(1)
    cp = theDoc.Paragraphs.count
    count = 1
    cs = 0
    For i = 1 To cp
        cs = 0
        If para.Range.Characters.count > 1 Then
            cs = para.Range.Sentences.count
        End If
        If cs <> 0 Then
            For j = 1 To cs
                Table1.Cell(count, number).Range.Text = para.Range.Sentences.item(j).Text
                Table1.Cell(count, number).Range.Find.Execute findtext:="^0013", replacewith:="", Replace:=wdReplaceAll
                If count = 100 Then
                    Exit For
                End If
                count = count + 1
            Next j
        End If
                If count = 100 Then
                    Exit For
                End If
'        Debug.Print khSent
        Set para = para.Next()
    Next i
End Sub
Sub ParallelToGlossed()
    Dim theDoc As Document
'    Set theDoc = Documents.Open("c:\ELAN\GAZETY\Elvira&Alya\Statii-tr-parall-1.doc") '"c:\ELAN\GAZETY\Parallel\statii-tr-parall.docx")
    Set theDoc = Application.ActiveDocument
    Dim newDoc As Document
    
    Dim actName As String
    actName = theDoc.Path & "\" & theDoc.name
    actName = Mid$(actName, 1, InStrRev(actName, ".") - 1)
    
    Set newDoc = Documents.Add()
    Set Range1 = newDoc.Range(start:=0, End:=0)
       
    Dim Table1 As table
    Set Table1 = theDoc.Tables.item(1)
    cr = Table1.Rows.count
    
    Dim newTable As table
    
    Dim punct As String
    punct = "-.,:;!?""'" & ChrW(&HBB) & ChrW(&HAB)

    Dim sizeSafeArray As Integer
    sizeSafeArray = 100
    
    Dim khPars As New khParser
    Dim res As Long
    
    res = khPars.Init(sizeSafeArray, "http://khakas.altaica.ru", actName & "_dict.txt", actName & "_notfound.txt")
'    res = khPars.Init(sizeSafeArray, "http://khakas.corpora.tk", actName & "_dict.txt", actName & "_notfound.txt")
    
    Dim arRows() As String
    Dim rgStr() As Integer ' SAFEARRAY
    ReDim rgStr(sizeSafeArray)
    
    If res = S_OK Then
        Dim c As Integer
        
        'из каждой строки таблицы будем делать отдельную таблицу
        'первая строка - предложение на хакасском
        'последняя строка - предложение на русском
        ' все остальные строки - слова + результаты парсирования
        Dim curCell As Cell
        Set curCell = Table1.Cell(1, 1)
        Dim newCell As Cell
        
        For i = 1 To cr
            c = curCell.Range.Words.count
            If c > 1 Then
                ReDim arRows(c)
                arRows(0) = curCell.Range.Text
                r = 1
                For j = 1 To c - 1
                    ch = Mid(curCell.Range.Words.item(j), 1, 1)
                    If InStr(punct, ch) = 0 Then
                        arRows(r) = curCell.Range.Words.item(j)
                        r = r + 1
                    End If
                Next j
                Set curCell = curCell.Next
                arRows(r) = curCell.Range.Text
                r = r + 1
            Else
                Set curCell = curCell.Next
            End If
            Set curCell = curCell.Next
            
'            AddTable = True
'            If AddTable Then
            Set newTable = newDoc.Tables.Add(Range1, r, 2)
            newTable.Borders.OutsideLineStyle = wdLineStyleSingle
            newTable.Borders.InsideLineStyle = wdLineStyleSingle
            newTable.AllowAutoFit = True
            
            newTable.Rows(1).Cells.Merge
            newTable.Rows(1).Range.Text = arRows(0)
            newTable.Rows(1).Range.Find.Execute findtext:="^0013", replacewith:="", Replace:=wdReplaceAll
            newTable.Rows(r).Cells.Merge
            newTable.Rows(r).Range.Text = arRows(r - 1)
            newTable.Rows(r).Range.Find.Execute findtext:="^0013", replacewith:="", Replace:=wdReplaceAll
            
            Set newCell = newTable.Cell(2, 1)
            For p = 2 To r - 1
                newCell.Range.Text = arRows(p - 1)
                Set newCell = newCell.Next
                Set newCell = newCell.Next
'                newTable.Cell(p, 1).Range.text = arRows(p - 1)
            Next p
            If newTable.Rows.count > 2 Then
                Set newCell = newTable.Cell(2, 2)
                For p = 1 To r - 2
                    res = khPars.DoParse(arRows(p))
                    If res = S_OK Then
                        numH = khPars.GetHomonymNum
                        If numH <> 0 Then
                            newCell.Split 2, numH
                            For cl = 1 To numH * 2
                                Hom = String(100, "*")
                                Call khPars.GetNextHomonym(rgStr)
                                Form = ""
                                For idx = 0 To sizeSafeArray
                                    ch = ChrW(rgStr(idx))
                                    If ch = "*" Then
                                        Exit For
                                    End If
                                    Form = Form + ch
                                Next idx
    
                                newCell.Range.Text = Form
                                Set newCell = newCell.Next
                            Next cl
                        Else
                            Set newCell = newCell.Next
                        End If
                        Set newCell = newCell.Next
                    End If
                Next p
            End If
            Selection.EndKey wdStory
            Selection.InsertAfter ("" & vbCrLf)
            Selection.EndKey wdStory
            Set Range1 = Selection.Range
'            Else ' AddTable = False
            For p = 0 To r - 1
                If p = 0 Then
                    res = khPars.AddKhakSent(Mid(arRows(p), 1, Len(arRows(p)) - 2))
                Else
                    If p = r - 1 Then
                        res = khPars.AddRusSent(Mid(arRows(p), 1, Len(arRows(p)) - 2))
                    Else
                        res = khPars.DoParse(arRows(p))
                    End If
                End If
            Next p
 '           End If ' AddTable
            Debug.Print i
            'If i = 300 Or i = 600 Or i = 900 Then
            '    Debug.Print i
            'End If
            
        Next i
    End If
'    Call khPars.SaveToELAN("c:\ELAN\GAZETY\Parallel\statii-tr-parall.eaf")
    Call khPars.SaveToELAN(actName & ".eaf")
    Call khPars.Terminate
    
End Sub

Sub ParallelToEAF()
    Set theDoc = Application.ActiveDocument
    Dim newDoc As Document
    
    Dim actName As String
    actName = theDoc.Path & "\" & theDoc.name
    actName = Mid$(actName, 1, InStrRev(actName, ".") - 1)
    
    Dim Table1 As table
    Set Table1 = theDoc.Tables.item(1)
    cr = Table1.Rows.count
    
    Dim newTable As table
    
    Dim punct As String
    punct = "-.,:;!?""'" & ChrW(&HBB) & ChrW(&HAB)

    Dim sizeSafeArray As Integer
    sizeSafeArray = 100
    
    Dim khPars As New khParser
    Dim res As Long
    res = khPars.Init(sizeSafeArray, "https://khakas.altaica.ru", actName & "_dict.txt", actName & "_notfound.txt")
    
    Dim arRows() As String
    Dim rgStr() As Integer ' SAFEARRAY
    ReDim rgStr(sizeSafeArray)
    
    If res = S_OK Then
        Dim c As Integer
        
        'из каждой строки таблицы будем делать отдельную таблицу
        'первая строка - предложение на хакасском
        'последняя строка - предложение на русском
        ' все остальные строки - слова + результаты парсирования
        Dim curCell As Cell
        Set curCell = Table1.Cell(1, 1)
        Dim newCell As Cell
        
        For i = 1 To cr
            c = curCell.Range.Words.count
            If c > 1 Then
                ReDim arRows(c)
                arRows(0) = curCell.Range.Text
                r = 1
                For j = 1 To c - 1
                    ch = Mid(curCell.Range.Words.item(j), 1, 1)
                    If InStr(punct, ch) = 0 Then
                        arRows(r) = curCell.Range.Words.item(j)
                        r = r + 1
                    End If
                Next j
                Set curCell = curCell.Next
                arRows(r) = curCell.Range.Text
                r = r + 1
            Else
                Set curCell = curCell.Next
            End If
            Set curCell = curCell.Next
            
            For p = 0 To r - 1
                If p = 0 Then
                    res = khPars.AddKhakSent(Mid(arRows(p), 1, Len(arRows(p)) - 2))
                Else
                    If p = r - 1 Then
                        res = khPars.AddRusSent(Mid(arRows(p), 1, Len(arRows(p)) - 2))
                    Else
                        res = khPars.DoParse(arRows(p))
                    End If
                End If
            Next p
        Next i
    End If
    
    Call khPars.SaveToELAN(actName & ".eaf")
    Call khPars.Terminate
    
End Sub
'делаем EAF-файл из таблицы, в которой размечены
'участники разговора
'в первом столбце таблицы - имена участников разговора
'в EAF-файле для каждого участника разговора будет создан свой слой
Sub ParallelToEAFWithName()
    Set theDoc = Application.ActiveDocument
    Dim newDoc As Document
    
    Dim actName As String
    actName = theDoc.Path & "\" & theDoc.name
    actName = Mid$(actName, 1, InStrRev(actName, ".") - 1)
    
    Dim Table1 As table
    Set Table1 = theDoc.Tables.item(1)
    cr = Table1.Rows.count
    
    Dim newTable As table
    
    Dim punct As String
    punct = "-.,:;!?""'" & ChrW(&HBB) & ChrW(&HAB)

    Dim sizeSafeArray As Integer
    sizeSafeArray = 100
    
    Dim khPars As New khParser
    Dim res As Long
    res = khPars.Init(sizeSafeArray, "https://khakas.altaica.ru", actName & "_dict.txt", actName & "_notfound.txt")
    
    Dim arRows() As String
    Dim rgStr() As Integer ' SAFEARRAY
    ReDim rgStr(sizeSafeArray)
    
    If res = S_OK Then
        Dim c As Integer
        
        Dim curCell As Cell
        Set curCell = Table1.Cell(1, 1)
'        Dim newCell As cell
        Dim name As String
        For i = 1 To cr
            name = Trim$(curCell.Range.Text)
            Set curCell = curCell.Next
            c = curCell.Range.Words.count
            If c > 1 Then
                ReDim arRows(c + 2)
                arRows(0) = name
                r = 1
                arRows(r) = curCell.Range.Text 'хакасское предложение
                r = r + 2
                For j = 1 To c - 1
                    ch = Mid(curCell.Range.Words.item(j), 1, 1)
                    If InStr(punct, ch) = 0 Then
                        arRows(r) = curCell.Range.Words.item(j)
                        r = r + 1
                    End If
                Next j
                Set curCell = curCell.Next
                arRows(r) = curCell.Range.Text 'русское предложение
                r = r + 1
            Else
                Set curCell = curCell.Next
            End If
            Set curCell = curCell.Next
            
            For p = 0 To r - 1
                If p = 0 Then
                    name = Mid(arRows(p), 1, Len(arRows(p)) - 2)
                Else
                    If p = 1 Then
                        res = khPars.AddKhakSent2(name, Mid(arRows(p), 1, Len(arRows(p)) - 2))
                    
                    Else
                        If p = r - 1 Then
                            res = khPars.AddRusSent(Mid(arRows(p), 1, Len(arRows(p)) - 2))
                        Else
                            res = khPars.DoParse(arRows(p))
                        End If
                    End If
                End If
            Next p
        Next i
    End If
    
    Call khPars.SaveToELAN(actName & ".eaf")
    Call khPars.Terminate
    
End Sub
Function getCellByName(name As String, rgSize As Integer, rgTitles() As String, rgIdx() As Integer, startCell As Cell, cc As Integer) As Cell
    Dim curCell As Cell
    Set curCell = startCell
    For i = 0 To rgSize
        If StrComp(name, rgTitles(i)) = 0 Then
            If rgIdx(i) >= 0 And rgIdx(i) <= cc - 1 Then
                For k = 0 To rgIdx(i) - 1
                    Set curCell = curCell.Next
                Next k
            Else
                Set curCell = Nothing
            End If
            Exit For
        End If
    Next i
    Set getCellByName = curCell
End Function
Function getName(rgTitles() As String, rgIdx() As Integer, startCell As Cell, cc As Integer) As Cell
    Set getName = getCellByName("name", 5, rgTitles, rgIdx, startCell, cc)
End Function
Function getTime(rgTitles() As String, rgIdx() As Integer, startCell As Cell, cc As Integer) As Cell
    Set getTime = getCellByName("time", 5, rgTitles, rgIdx, startCell, cc)
End Function
Function getTranscription(rgTitles() As String, rgIdx() As Integer, startCell As Cell, cc As Integer) As Cell
    Set getTranscription = getCellByName("transcription", 5, rgTitles, rgIdx, startCell, cc)
End Function
Function getKhak(rgTitles() As String, rgIdx() As Integer, startCell As Cell, cc As Integer) As Cell
    Set getKhak = getCellByName("khak", 5, rgTitles, rgIdx, startCell, cc)
End Function
Function getRus(rgTitles() As String, rgIdx() As Integer, startCell As Cell, cc As Integer) As Cell
    Set getRus = getCellByName("rus", 5, rgTitles, rgIdx, startCell, cc)
End Function
Function nextRowCell(startCell As Cell, cc As Integer) As Cell
    Dim curCell As Cell
    Set curCell = startCell
    For i = 0 To cc - 1
        Set curCell = curCell.Next
    Next i
    Set nextRowCell = curCell
End Function

'делаем EAF-файл из таблицы, в которой размечен текст с привязкой по времени
'(информант) | (время начала предложения, например, 00.02.54) | хакасский текст | русский текст
'в EAF-файле для каждого участника разговора будет создан свой слой
Sub ParallelToEAFWithNameAndTime()
    Set theDoc = Application.ActiveDocument
    Dim newDoc As Document
    
    Dim actName As String
    actName = theDoc.Path & "\" & theDoc.name
    actName = Mid$(actName, 1, InStrRev(actName, ".") - 1)
    
    Dim Table1 As table
    Set Table1 = theDoc.Tables.item(1)
    cr = Table1.Rows.count
    
    Dim newTable As table
    
    Dim punct As String
    punct = "-.,:;!?""'" & ChrW(&HBB) & ChrW(&HAB)

    Dim sizeSafeArray As Integer
    sizeSafeArray = 100
    
    Dim khPars As New khParser
    Dim res As Long
    res = khPars.Init(sizeSafeArray, "https://khakas.altaica.ru", actName & "_dict.txt", actName & "_notfound.txt")
    
    Dim arRows() As String
    Dim rgStr() As Integer ' SAFEARRAY
    ReDim rgStr(sizeSafeArray)
        
    Dim rgTitles(5) As String
    rgTitles(0) = "name"
    rgTitles(1) = "time"
    rgTitles(2) = "transcription"
    rgTitles(3) = "khak"
    rgTitles(4) = "rus"
        
    Dim rgIdx(5) As Integer
    
    If res = S_OK Then
        Dim c As Integer
        
        Dim curCell As Cell
        Set curCell = Table1.Cell(1, 1)
        Dim cc As Integer
        cc = Table1.Rows.item(1).Cells.count
        For j = 0 To 4
            rgIdx(j) = -1
            Set curCell = Table1.Cell(1, 1)
            For i = 0 To cc - 1
                txt = Trim$(curCell.Range.Text)
                txt = Mid(txt, 1, Len(txt) - 2)
                If StrComp(txt, rgTitles(j)) = 0 Then
                    rgIdx(j) = i
                    Exit For
                End If
                Set curCell = curCell.Next
            Next i
        Next j
        Dim name As String
        Dim time As String
        Dim startCell As Cell
        Set startCell = nextRowCell(Table1.Cell(1, 1), cc)
        For i = 2 To cr
            name = ""
            Set curCell = getName(rgTitles, rgIdx, startCell, cc)
            If Not curCell Is Nothing Then
                name = Trim$(curCell.Range.Text)
            End If
            time = ""
            Set curCell = getTime(rgTitles, rgIdx, startCell, cc)
            If Not curCell Is Nothing Then
                time = Trim$(curCell.Range.Text)
            End If
            Set curCell = getKhak(rgTitles, rgIdx, startCell, cc)
            c = 0
            khak = ""
            If Not curCell Is Nothing Then
                khak = curCell.Range.Text 'хакасское предложение
                c = curCell.Range.Words.count
            End If
            ReDim arRows(c + 5)
            arRows(0) = name
            arRows(1) = time
            arRows(2) = khak
            If c > 1 Then
                r = 3
                For j = 1 To c - 1
                    ch = Mid(curCell.Range.Words.item(j), 1, 1)
                    If InStr(punct, ch) = 0 Then
                        arRows(r) = curCell.Range.Words.item(j)
                        r = r + 1
                    End If
                Next j
            Else
                r = 3
            End If
            rus = ""
            Set curCell = getRus(rgTitles, rgIdx, startCell, cc) 'curCell.Next
            If Not curCell Is Nothing Then
                rus = curCell.Range.Text
            End If
            arRows(r) = rus 'русское предложение
            r = r + 1
            transcr = ""
            Set curCell = getTranscription(rgTitles, rgIdx, startCell, cc) 'curCell.Next
            If Not curCell Is Nothing Then
                transcr = curCell.Range.Text
            End If
            arRows(r) = transcr 'транскрипция
            r = r + 1
            Set startCell = nextRowCell(startCell, cc) 'curCell.Next
            
            name = ""
            time = ""
            khak = ""
            transcr = ""
            rus = ""
            For p = 0 To r - 1
                If p = 0 And Len(arRows(p)) > 2 Then
                    name = Mid(arRows(p), 1, Len(arRows(p)) - 2)
                ElseIf p = 1 And Len(arRows(p)) > 2 Then
                    time = Mid(arRows(p), 1, Len(arRows(p)) - 2)
                ElseIf p = 2 Then
                    If Len(arRows(p)) > 2 Then
                        khak = Mid(arRows(p), 1, Len(arRows(p)) - 2)
                    End If
                    res = khPars.AddKhakSent3(name, time, khak)
                ElseIf p = r - 2 Then
                    If Len(arRows(p)) > 2 Then
                        res = khPars.AddRusSent(Mid(arRows(p), 1, Len(arRows(p)) - 2))
                    Else
                        res = khPars.AddRusSent(rus)
                    End If
                ElseIf p = r - 1 Then
                    If Len(arRows(p)) > 2 Then
                        res = khPars.AddTranscription(Mid(arRows(p), 1, Len(arRows(p)) - 2))
                    Else
                        res = khPars.AddTranscription(transcr)
                    End If
                Else
                    res = khPars.DoParse(arRows(p))
                End If
            Next p
        Next i
    End If
    
    Call khPars.SaveToELANFlexTime(actName & ".eaf")
    Call khPars.Terminate
    
End Sub


