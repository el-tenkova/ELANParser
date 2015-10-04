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
                Table1.Cell(count, number).Range.text = para.Range.Sentences.item(j).text
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
Sub MakeParallel()
    Dim rusDoc As Document
    Dim khakDoc As Document
    Dim newDoc As Document
    
    Set rusDoc = Documents.Open("c:\ELAN\GAZETY\stati-troshkina-ost.docx")
    Set khakDoc = Documents.Open("c:\ELAN\GAZETY\khak-stati-troshkina-ost.docx")
    
    Set newDoc = Documents.Add()
    Set Range1 = newDoc.Range(start:=0, End:=0)
   
    'Dim paraR As Paragraph
    'Dim paraKH As Paragraph
    Dim Table1 As table
    
    Dim khSent As Long
    khSent = getSentCount(khakDoc)
    Dim rSent As Long
    rSent = getSentCount(rusDoc)
    
    Dim rc As Long
    rc = khSent
    If rc < rSent Then
        rc = rSent
    End If
    khSent = 100
    rSent = 100
    
    Set Table1 = newDoc.Tables.Add(Range1, rc, 2)
    Table1.Borders.OutsideLineStyle = wdLineStyleSingle
    Table1.Borders.InsideLineStyle = wdLineStyleSingle
   
    Call writeColumn(khakDoc, 1, Table1)
    Call writeColumn(rusDoc, 2, Table1)
    
End Sub

Sub ParallelToELAN()
    Dim i As Long
    Dim j As Long
    Dim c As Long

    Dim Title As String
    Dim begin As LongLong

    Dim theDoc As Document
    Set theDoc = Documents.Open("c:\ELAN\GAZETY\elantest.docx")
    
    Dim Table1 As table
    Set Table1 = theDoc.Tables.item(1)
    cr = Table1.Rows.count
       
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open

'Dim esT As Object
'Set esT = CreateObject("ADODB.Stream")
'esT.Type = 2
'esT.Charset = "utf-8"
'esT.Open

    Dim elProc As New ELANProc
    
't = ActiveDocument.Tables.count
    ' «апись заголовка файла
    Call elProc.writeHeader(fsT)
    ' «апись временных интервалов
    fsT.WriteText ("<TIME_ORDER>" & vbCrLf)
   ' cells_cnt = 0
    begin = 0
    Dim time_intervals() As LongLong
    ReDim time_intervals(1 To 2 * (cr + 1))
    time_intervals(1) = 1
    j = 1
    begin = 0
    j = 1
    For i = 1 To cr + 1
'        Set tabl = ActiveDocument.Tables.item(i)
'        If tabl.Rows.count = 4 Then
'        time_intervals(j + 1) = cells_cnt + tabl.Range.Cells.count + 1
 '       cells_cnt = cells_cnt + tabl.Range.Cells.count
        time_intervals(j) = j
        Call elProc.writeTimeSlot(j, begin, fsT)
        j = j + 1
        time_intervals(j) = j
        Call elProc.writeTimeSlot(j, begin, fsT)
        j = j + 1
        
        begin = begin + 500
'    End If
    Next i
'Call writeTimeSlot(time_intervals(j), time_intervals(j) + 4, begin, fsT)
    fsT.WriteText ("</TIME_ORDER>" & vbCrLf)

    Dim row As Object
    Dim annoID As Long
    annoID = 1
    Dim delta As Integer
    ' «апись слоЄв
    Dim time1 As LongLong
    Dim time2 As LongLong
    Dim text As String
    For Num = 1 To 2
        fsT.WriteText ("<TIER DEFAULT_LOCALE=""ru"" LINGUISTIC_TYPE_REF=""default"" TIER_ID=""Tier-")
        fsT.WriteText (Format$(Num - 1) & """")
        fsT.WriteText (">" & vbCrLf)
        delta = 2
        'timeJ = 1
'        For i = 1 To cr
'            Set tabl = ActiveDocument.Tables.item(i)
'        If tabl.Rows.count = 4 Then
'            timeIdx = time_intervals(timeJ) + delta
'            Set row = tabl.Rows.item(Num)
        timeIdx = Num
        For j = 1 To cr 'row.Cells.count
            text = Table1.Cell(j, Num).Range.text
'                row.Cells.item (j)
            time1 = timeIdx
            time2 = timeIdx + delta
                
            'If Num = 4 Then
            '        time2 = time_intervals(timeJ + 1)
            '    Else
            '        If j = 1 Then
            '            time2 = time2 + 1
            '        End If
            '    End If
            Call elProc.writeAnno(Mid$(text, 1, Len(text) - 2), annoID, time1, time2, fsT)
            annoID = annoID + 1
            timeIdx = time2
        Next j
            'timeJ = timeJ + 1
        'Else
        '    esT.WriteText ("¬ таблице номер " & Format$(i) & " меньше четырЄх строк" & vbCrLf)
            'esT.WriteText (Format$(timeIdx))
        'End If
        'Next i
        fsT.WriteText ("</TIER>" & vbCrLf)
    Next Num
    ' «апись последнего блока файла
    Call elProc.writeTail(fsT)
    NameCSV = theDoc.Path & "\" & Mid$(theDoc.Name, 1, InStrRev(theDoc.Name, ".")) & "eaf"
    fsT.SaveToFile NameCSV, 2
    fsT.Close
    'esT.SaveToFile ActiveDocument.Path & "\!!err.txt", 2
    'esT.Close

End Sub

Sub ParallelToGlossed()
    Dim theDoc As Document
    Set theDoc = Documents.Open("c:\ELAN\GAZETY\Parallel\statii-tr-parall.docx")
    Dim newDoc As Document
    
    Set newDoc = Documents.Add()
    Set Range1 = newDoc.Range(start:=0, End:=0)
       
    Dim Table1 As table
    Set Table1 = theDoc.Tables.item(1)
    cr = 5 'Table1.Rows.count
    
    Dim newTable As table
    
    Dim punct As String
    punct = "-.,:;!?""'" & ChrW(&HBB) & ChrW(&HAB)

    Dim sizeSafeArray As Integer
    sizeSafeArray = 100
    
    Dim khPars As New khParser
    Dim res As Long
    res = khPars.Init(sizeSafeArray, "c:\ELAN\GAZETY\Parallel\dict-st-tr.txt", "c:\ELAN\GAZETY\Parallel\notfound-st-tr.txt")
    
    Dim arRows() As String
    Dim rgStr() As Integer ' SAFEARRAY
    ReDim rgStr(sizeSafeArray)
    
    If res = S_OK Then
        Dim c As Integer
        
        'из каждой строки таблицы будем делать отдельную таблицу
        'перва€ строка - предложение на хакасском
        'последн€€ строка - предложение на русском
        ' все остальные строки - слова + результаты парсировани€
        Dim curCell As Cell
        Set curCell = Table1.Cell(1, 1)
        Dim newCell As Cell
        
        For i = 1 To cr + 1
            c = curCell.Range.Words.count
            If c > 1 Then
                ReDim arRows(c)
                arRows(0) = curCell.Range.text
                r = 1
                For j = 1 To c - 1
                    ch = Mid(curCell.Range.Words.item(j), 1, 1)
                    If InStr(punct, ch) = 0 Then
                        arRows(r) = curCell.Range.Words.item(j)
                        r = r + 1
                    End If
                Next j
                Set curCell = curCell.Next
                arRows(r) = curCell.Range.text
                r = r + 1
            Else
                Set curCell = curCell.Next
            End If
            Set curCell = curCell.Next
            
            AddTable = False
            If AddTable Then
            Set newTable = newDoc.Tables.Add(Range1, r, 2)
            newTable.Borders.OutsideLineStyle = wdLineStyleSingle
            newTable.Borders.InsideLineStyle = wdLineStyleSingle
            newTable.AllowAutoFit = True
            
            newTable.Rows(1).Cells.Merge
            newTable.Rows(1).Range.text = arRows(0)
            newTable.Rows(1).Range.Find.Execute findtext:="^0013", replacewith:="", Replace:=wdReplaceAll
            newTable.Rows(r).Cells.Merge
            newTable.Rows(r).Range.text = arRows(r - 1)
            newTable.Rows(r).Range.Find.Execute findtext:="^0013", replacewith:="", Replace:=wdReplaceAll
            
            Set newCell = newTable.Cell(2, 1)
            For p = 2 To r - 1
                newCell.Range.text = arRows(p - 1)
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
    
                                newCell.Range.text = Form
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
            Else ' AddTable = False
            For p = 1 To r - 2
                res = khPars.DoParse(arRows(p))
            Next p
            End If ' AddTable
            Debug.Print i
            'If i = 300 Or i = 600 Or i = 900 Then
            '    Debug.Print i
            'End If
            
        Next i
    End If
    Call khPars.Terminate
    
End Sub
Sub ParallelToUTF8()
    Dim theDoc As Document
    Set theDoc = Documents.Open("c:\ELAN\GAZETY\Parallel\statii-tr-parall.docx")
    
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    
    Dim Table1 As table
    Set Table1 = theDoc.Tables.item(1)
    cr = Table1.Rows.count
    
'    Dim khParser As New KhakasParser
'    Call khParser.Init
    
    Dim pars As Object
    Set pars = New khParser
    
    Call pars.Init
    
    
    
    Dim punct As String
    punct = ".,:;!?""'" & ChrW(&HBB) ' & ChrW(&HAB)

    Dim c As Integer
    'из каждой строки таблицы будем делать блок из строк в файле
    For i = 1 To cr + 1
        c = Table1.Cell(i, 1).Range.Words.count
        If c > 1 Then
            col = 1
            doMerge = False
            fsT.WriteText ("-----------------------------------------------" & vbCrLf)
            ' перва€ строка - хакасский текст
            fsT.WriteText ("KHAK: " & Mid$(Table1.Cell(i, 1).Range.text, 1, Len(Table1.Cell(i, 1).Range.text) - 1) & vbCrLf)
            kav = False
            For j = 1 To c - 1
                ' кажда€ последующа€ строка - хакасское слово в исходном виде, потом через | в нормализованном виде
                ch = Mid(Table1.Cell(i, 1).Range.Words.item(j), 1, 1)
                If InStr(punct, ch) = 0 Then
                    If kav = False And j <> 1 Then
                        fsT.WriteText (vbCrLf)
                    End If
                    If ch = ChrW(&HAB) Then
                        kav = True
                        fsT.WriteText ("KHAK: " & Table1.Cell(i, 1).Range.Words.item(j))
                    Else
                        If kav = True Then
                            fsT.WriteText (Table1.Cell(i, 1).Range.Words.item(j))
                            fsT.WriteText ("|" & khParser.Normalize(Table1.Cell(i, 1).Range.Words.item(j)))
                            kav = False
                        Else
                            fsT.WriteText ("KHAK: " & Table1.Cell(i, 1).Range.Words.item(j))
                            fsT.WriteText ("|" & khParser.Normalize(Table1.Cell(i, 1).Range.Words.item(j)))
                        End If
                    End If
                Else
                    fsT.WriteText (Table1.Cell(i, 1).Range.Words.item(j))
                End If
            Next j
            ' последн€€ строка - русский текст
            fsT.WriteText (vbCrLf & "RUS: " & Mid$(Table1.Cell(i, 2).Range.text, 1, Len(Table1.Cell(i, 2).Range.text) - 1) & vbCrLf)
        End If
        Debug.Print i
    Next i
    NameUTF8 = theDoc.Path & "\" & Mid$(theDoc.Name, 1, InStrRev(theDoc.Name, ".")) & "txt"
    fsT.SaveToFile NameUTF8, 2
    fsT.Close
End Sub
