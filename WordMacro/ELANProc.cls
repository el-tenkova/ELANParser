VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ELANProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function writeHeader(fileStream As Object)
    header = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    header = header & "<ANNOTATION_DOCUMENT AUTHOR=""unspecified"" DATE=""2015-04-09T12:42:10+04:00"" FORMAT=""2.8"" VERSION=""2.8"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""http://www.mpi.nl/tools/elan/EAFv2.8.xsd"">" & vbCrLf
    header = header & "<HEADER MEDIA_FILE="""" TIME_UNITS=""milliseconds"">" & vbCrLf
    header = header & "<MEDIA_DESCRIPTOR MEDIA_URL="""" MIME_TYPE=""audio/x-wav""/>" & vbCrLf 'to avoid Annex reject error
        '<PROPERTY NAME="URN">urn:nl-mpi-tools-elan-eaf:6504c785-9396-4c0d-abc5-c224bfd7eb09</PROPERTY>
        '<PROPERTY NAME="lastUsedAnnotationId">49</PROPERTY>
    header = header & "</HEADER>" & vbCrLf
    
    fileStream.WriteText header
End Function
Public Function writeTail(fileStream As Object)
    tail = "<LINGUISTIC_TYPE GRAPHIC_REFERENCES=""false"" LINGUISTIC_TYPE_ID=""default"" TIME_ALIGNABLE=""true""/>" & vbCrLf
    tail = tail & "<LOCALE COUNTRY_CODE=""EN"" LANGUAGE_CODE=""us""/>" & vbCrLf
    tail = tail & "<LOCALE COUNTRY_CODE=""RU"" LANGUAGE_CODE=""ru""/>""" & vbCrLf
    tail = tail & "<CONSTRAINT DESCRIPTION=""Time subdivision of parent annotation's time interval, no time gaps allowed within this interval"" STEREOTYPE=""Time_Subdivision""/>" & vbCrLf
    tail = tail & "<CONSTRAINT DESCRIPTION=""Symbolic subdivision of a parent annotation. Annotations refering to the same parent are ordered"" STEREOTYPE=""Symbolic_Subdivision""/>" & vbCrLf
    tail = tail & "<CONSTRAINT DESCRIPTION=""1-1 association with a parent annotation"" STEREOTYPE=""Symbolic_Association""/>" & vbCrLf
    tail = tail & "<CONSTRAINT DESCRIPTION=""Time alignable annotations within the parent annotation's time interval, gaps are allowed"" STEREOTYPE=""Included_In""/>" & vbCrLf
    tail = tail & "</ANNOTATION_DOCUMENT>" & vbCrLf
    fileStream.WriteText tail
End Function

Public Function writeTimeSlotEx(idx1 As Integer, idx2 As Integer, ByRef begin As LongLong, fileStream As Object)
    c = (idx2 - idx1) / 3
    p = 1
    j = idx1
    For p = 1 To c
        For k = 1 To 3
            fileStream.WriteText ("<TIME_SLOT TIME_SLOT_ID=""ts" & Format$(j) & """ TIME_VALUE=")
            fileStream.WriteText ("""" & Format$(begin) & """/>" & vbCrLf)
            j = j + 1
        Next k
        If p = 1 Then
            fileStream.WriteText ("<TIME_SLOT TIME_SLOT_ID=""ts" & Format$(j) & """ TIME_VALUE=")
            fileStream.WriteText ("""" & Format$(begin) & """/>" & vbCrLf)
            j = j + 1
        End If
        begin = begin + 500
    Next p
End Function
Public Function writeAnno(text As String, id As Long, time1 As LongLong, time2 As LongLong, fileStream As Object)
    text = Replace(text, "&H001f", "")
    txt = "<ANNOTATION>" & vbCrLf
    txt = txt & "<ALIGNABLE_ANNOTATION ANNOTATION_ID="""
    txt = txt & "a" & Format$(id) & """"
    txt = txt & " TIME_SLOT_REF1="""
    txt = txt & "ts" & Format$(time1) & """"
    txt = txt & " TIME_SLOT_REF2="""
    txt = txt & "ts" & Format$(time2) & """"
    txt = txt & ">" & vbCrLf
    txt = txt & "<ANNOTATION_VALUE>" & text & "</ANNOTATION_VALUE>" & vbCrLf
    txt = txt & "</ALIGNABLE_ANNOTATION>" & vbCrLf
    txt = txt & "</ANNOTATION>" & vbCrLf
    
    fileStream.WriteText (txt)
End Function

Public Function writeTimeSlot(idx As Long, begin As LongLong, fileStream As Object)
    fileStream.WriteText ("<TIME_SLOT TIME_SLOT_ID=""ts" & Format$(idx) & """ TIME_VALUE=")
    fileStream.WriteText ("""" & Format$(begin) & """/>" & vbCrLf)
End Function


