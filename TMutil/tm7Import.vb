Imports System.Xml.Serialization
Imports System.IO
Imports System.Data

Public Class tm7Import
    Public xD As XDocument
    Public Sub New(tm7file$)
        If Dir(tm7file$) = "" Then
            Console.WriteLine("File not found - exiting")
            Exit Sub
        End If

        '        xmlDataSet = New System.Data.DataSet 'DataSet
        '        Dim xmlSerializer As XmlSerializer = New XmlSerializer(xmlDataSet.GetType)
        ' Dim readStream As FileStream = New FileStream(tm7file, FileMode.Open)
        '        xmlDataSet = CType(xmlSerializer.Deserialize(readStream), System.Data.DataSet)
        ' readStream.Close()

        ' xD = XDocument.Load(tm7file)

        Dim threatCats$ = "<a:ThreatType>,UserThreatCategory,UserThreatDescription,UserThreatShortDescription,ThreatCategories,ThreatCategory,ThreatMetadata,ThreatTypes"
        Dim finD As Collection = CSVtoCOLL(threatCats)

        Dim tm7Str$ = streamReaderTxt(tm7file)
        Dim ctR As Long = 1
        ' Dim maxCtr As Long = Len(tm7Str)

        Dim parsE$ = tm7Str

        '   For Each F In finD
        '   ctR = 1
        '   parsE = tm7Str
        '         If InStr(F, "ThreatType>") = 0 Then GoTo nextFind
        Dim F$ = "<a:ThreatType>"

        Do Until InStr(parsE, F) = 0
            Dim endAt As Integer = 0
            ctR = InStr(parsE, F)
            endAt = InStr(parsE, "</a:ThreatType>") + 15
            Dim tType$ = Mid(parsE, ctR, endAt - ctR)
            Console.WriteLine(F + ": " + Replace(tType, "</", vbCrLf + "</") + "--------------------------------" + vbCrLf)
            parsE = Mid(parsE, endAt)
        Loop
nextFind:
        'Next

        Dim k As Integer
        k = 1

    End Sub
End Class
