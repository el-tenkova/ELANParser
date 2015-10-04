VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KhakasParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private objHTTP As Object
'Private strServ As String
'Private strData As String
'Private regExpObj As Object
'Private replCollection As Object
'Private homonyms As Collection
Private objKhParser As Object
Private dict As String
Private notFound As String
Public Function Init(sizeSafeArray As Long, dictPath As String, notfoundPath As String) As Long
    Set objKhParser = New khParser
    res = objKhParser.Init(sizeSafeArray, dictPath, notfoundPath)
'    strServ = "http://khakas.altaica.ru/cgi-bin/suddenly.fs?Getparam=getval"
'    strData = "&parse="

'    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
'    Set regExpObj = createObjToNormalize()
'    Set replCollection = createReplToNormalize()
    Init = res
End Function
'Public Function Normalize(inWord As String) As String
'    Dim str As String
    
'    str = LCase$(inWord)
'    Set colMatches = regExpObj.Execute(str) ' Execute search.
'    For Each objMatch In colMatches   ' Iterate Matches collection.
'        str = Replace$(str, objMatch, replCollection.item(objMatch))
'    Next objMatch
'    Normalize = str

'End Function


Public Function DoParse(word As String) As Integer

'    Dim normWord As String
    
'    normWord = Normalize(word)

    Call objKhParser.DoParse(word)
'    objHTTP.Open "post", strServ & strData & normWord, False
'    objHTTP.Send strData
'    If objHTTP.Status = 200 Then
'        Set homonyms = ParseResponse()
'    End If
'    DoParse = objHTTP.Status

End Function
'Function getDetails(str As String) As String
'    Dim aff As String
'    aff = ""
'    dash = False
'    str = Trim$(str)
'    For i = 1 To Len(str)
'        ch = Mid(str, i, 1)
'        If ch <> "-" Then
'            If dash = True Then
'                aff = aff & "-"
'            End If
'            aff = aff & ch
'            dash = False
'        Else
'            dash = True
'        End If
'    Next i
'    getDetails = aff
'End Function
'Public Function getMorphems() As String
'    Dim morphems As String
'    morphems = ""
'    For Each homonym In homonyms   ' Iterate homonyms
'        If Len(morphems) Then
'             morphems = morphems & vbCrLf
'        End If
'        If homonym.item("has-stem") = "1" Then
'            morphems = morphems & " " & Trim(homonym.item("stem"))
'        Else
'            morphems = morphems & Trim(homonym.item("headword"))
'        End If
'        morphems = morphems & getDetails(homonym.item("affixes"))
'
'    Next homonym
'    getMorphems = morphems
'
'End Function
'Public Function getGrammatics() As String
'    Dim gramm As String
'    gramm = ""
'    For Each homonym In homonyms   ' Iterate homonyms
'        If Len(gramm) Then
'             gramm = gramm & vbCrLf
'        End If
'        meaning = Trim$(homonym.item("meaning"))
'        meaning = Mid$(meaning, 2, Len(meaning) - 2)
'        gramm = gramm & meaning
'        gramm = gramm & getDetails(homonym.item("form"))
'
'    Next homonym
'    getGrammatics = gramm
'
'End Function
'Public Function ParseResponse() As Collection
'    Dim i As Long
'    Dim j As Long
'    Dim k As Long
'    Dim str As String
'    Dim result As New Collection
'    Dim idx As Long
'    idx = 1
'
'    str = objHTTP.responsetext
'
'    i = InStr(str, "FOUND STEM:")
'    Dim homonym As Collection
'    Do While i <> 0
'        Set homonym = New Collection
'        str = Mid(str, i + 12)
'        j = InStr(str, " ")
'        i = InStr(str, Chr(10))
'        'Debug.Print str
        
'        homonym.Add Mid(str, 1, j), "form"
'        homonym.Add Mid(str, j + 1, i - (j + 1)), "affixes"
'        str = Mid(str, i + 1)
'        homonym.Add Mid(str, 1, 1), "p_o_s"
'        i = InStr(str, Chr(10))
'
'        dict = Mid(str, 2, i - 2)
'        j = InStr(dict, " " & ChrW(&H201B))
'        homonym.Add Mid(dict, 1, j), "headword"
'        k = InStr(dict, ChrW(&H2019))
'        homonym.Add Mid(dict, j + 1, k + 1 - (j + 1)), "meaning"
'
'        homonym.Add "0", "has-stem"
'        If (Len(dict) > k + 1) Then
'            homonym.Add Mid(dict, k + 2), "stem"
'            homonym.Remove "has-stem"
'            homonym.Add "1", "has-stem"
            
'        End If
'
'        result.Add homonym, CStr(idx)
'        idx = idx + 1
        
'        i = InStr(str, "FOUND STEM:")
'    Loop
'    Set ParseResponse = result
    
'End Function

'Private Function createObjToNormalize() As Object
    ' Create a regular expression object.
'    Set objRegExp = New RegExp

    'Set the pattern by using the Pattern property.
'    objRegExp.Pattern = "[aceiopxy" & ChrW$(&HF6) & ChrW$(&HFF) & ChrW$(&H4B7) & "]"

    ' Set Case Insensitivity.
'    objRegExp.IgnoreCase = True

    'Set global applicability.
'    objRegExp.Global = True

    'Test whether the String can be compared.
'     Set createObjToNormalize = objRegExp
    
'End Function
'Private Function createReplToNormalize() As Collection
'    Dim repl As New Collection
'    repl.Add "à", "a"
'    repl.Add "ñ", "c"
'    repl.Add "å", "e"
'    repl.Add ChrW$(&H456), "i"
'    repl.Add "î", "o"
'    repl.Add "ð", "p"
'    repl.Add "õ", "x"
'    repl.Add "ó", "y"
'    repl.Add ChrW$(&H4E7), ChrW$(&HF6)
'    repl.Add ChrW$(&H4F1), ChrW$(&HFF)
'    repl.Add ChrW$(&H4CC), ChrW$(&H4B7)
''   'a': 'à', 'c': 'ñ', 'e': 'å', 'i': '\u0456', 'o': 'î', 'p': 'ð', 'x': 'õ', 'y': 'ó', '\u00F6': '\u04E7', '\u00FF': '\u04F1', '\u04B7': '\u04CC'
'    Set createReplToNormalize = repl
    
'End Function

