Option Explicit

Private http As Object

Private Sub Class_Initialize()
    Set http = CreateObject("MSXML2.XMLHTTP.3.0")
End Sub

Private Sub Class_Terminate()
    Set http = Nothing
End Sub

Public Function SZSE_Profile(Optional ByVal query As String = "0", Optional ByVal limit As Integer = 5000) As Dicts
    
    Dim d As New Dicts
    Dim i, j
    
    With http
        .Open "POST", "http://xbrl.cninfo.com.cn/do/stockreserch/getcompanybyprefix", False
        .setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .send "ticker=" & query & "&limit=" & limit & "&date=365"
        
        For Each i In Split(.responseText, ",")
            j = Split(i, "#")
            d.dict("SZ" & format(j(0), "000000")) = j
        Next i
    End With
    
    Set SZSE_Profile = d
    Set d = Nothing

End Function


Public Function SZSE_GeneralInfo(Optional ByVal id As String = "1") As Dicts
    
    Dim d As New Dicts
    ' '* Tools->Refernces Microsoft HTML Object Library
    Dim doc As MSHTML.HTMLDocument
    Dim i
    Dim j As Object
    
    Dim l As New Lists
    Dim l1 As New Lists
    
    id = format(id, "000000")
    
    With http
        .Open "POST", "http://xbrl.cninfo.com.cn/do/generalinfo/getcompanygeneralinfo", False
        .setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .send "ticker=" & id
        
        If .readyState = 4 And .Status = 200 Then
            Set doc = New MSHTML.HTMLDocument
            doc.body.innerHTML = .responseText
            
            l.add l1.fromArray(Split(doc.querySelector("#companyGenerationInfo_table4").innerText, Chr(13)))

            Set j = doc.querySelectorAll(".companyGenerationInfo_table6")
            For i = 0 To j.length - 1
                l.add l1.fromArray(Split(j.Item(i).innerText, Chr(13)))
            Next i
            
            With l.zipMe
                d.dict = d.arrToDict(l1.fromArray(Split(doc.querySelector("#companyGenerationInfo_table3").innerText, Chr(13))).setVal(0, "SE" & id & "-" & SZSE_Profile(id, 1).dict("SZ" & id)(1)).toArray, .toArray)
                d.label = l.getVal(0)
            End With
        Else
            MsgBox "Error" & vbNewLine & "Ready state: " & .readyState & _
            vbNewLine & "HTTP request status: " & .Status
        End If
    End With

    Set SZSE_GeneralInfo = d
    Set d = Nothing
    Set j = Nothing
    Set l = Nothing
    Set l1 = Nothing
    Set doc = Nothing

End Function

Public Function SZSE_Punishment(Optional ByVal id As String = "1") As Dicts
    
    Dim d As New Dicts
    ' '* Tools->Refernces Microsoft HTML Object Library
    Dim doc As MSHTML.HTMLDocument
    Dim i
    Dim j As Object
    
    Dim l As New Lists
    Dim l1 As New Lists
    
    Dim total As Integer
    Dim cnt As Integer
    cnt = 1
    
    id = format(id, "000000")
    
    Set doc = post("http://xbrl.cninfo.com.cn/do/sincerelycase/getpunishmentdate", "ticker=" & id & "&page=" & cnt)
    total = CInt(doc.querySelector("#pageCount").getAttribute("value"))
    Set l = l.fromArray(Split(doc.querySelector("#notice").innerText, Chr(13)))
    
    Do While cnt < total
        cnt = cnt + 1
        Set doc = post("http://xbrl.cninfo.com.cn/do/sincerelycase/getpunishmentdate", "ticker=" & id & "&page=" & cnt)
        l.addAll (l1.fromArray(Split(doc.querySelector("#notice").innerText, Chr(13))))
    Loop
    
    ' againfinance_table
    For Each i In l.toArray
        Set doc = post("http://xbrl.cninfo.com.cn/do/sincerelycase/getsincerelycase", "ticker=" & id & "&date=" & i & "&index=0")
        l1.add domToList(".sincerelycaseDetail_td1", doc)
    Next i

    d.dict = d.arrToDict(l.map("""D_""&""_""").toArray, l1.toArray)
    Set SZSE_Punishment = d
    
    Set d = Nothing
    Set j = Nothing
    Set l = Nothing
    Set l1 = Nothing
    Set doc = Nothing
    
End Function


Public Function SZSE_Dividend(Optional ByVal id As String = "1") As Dicts
    
    Dim d As New Dicts
    ' '* Tools->Refernces Microsoft HTML Object Library
    Dim doc As MSHTML.HTMLDocument
    Dim i
    Dim j As Object
    
    Dim l As New Lists
    Dim l1 As New Lists
    
    Dim total As Integer
    Dim cnt As Integer
    cnt = 1
    
    id = format(id, "000000")
    
    
    Set doc = post("http://xbrl.cninfo.com.cn/do/dividend/getdividendhistory", "ticker=" & id & "&page=" & cnt)
    total = CInt(doc.querySelector("#pageCount").getAttribute("value"))
    Set l = domToList(".dividendHistorySubpage_table tr td", doc)
    
    Do While cnt < total
        cnt = cnt + 1
        Set doc = post("http://xbrl.cninfo.com.cn/do/dividend/getdividendhistory", "ticker=" & id & "&page=" & cnt)
        l.addAll domToList(".dividendHistorySubpage_table tr td", doc).drop(5)
        
        Debug.Print cnt
    Loop
    
    Set l = l.subgroupBy(5, 5)
    
    With l
        d.dict = d.arrToDict(.zipMe.getVal(0).map("""D*""&""_""").toArray, .toArray)
        d.label = l.getVal(0)
    End With
    
    d.p
    
    Set SZSE_Dividend = d
    
    Set d = Nothing
    Set j = Nothing
    Set l = Nothing
    Set l1 = Nothing
    Set doc = Nothing
    
End Function

Private Function post(ByVal url As String, Optional ByVal data As String) As MSHTML.HTMLDocument
    
    Dim doc As MSHTML.HTMLDocument
    
    With http
        .Open "POST", url, False
        .setRequestHeader "content-type", "application/x-www-form-urlencoded"
        If IsMissing(data) Or data = "" Then
            .send
        Else
            .send data
        End If
        
        If .readyState = 4 And .Status = 200 Then
            Set doc = New MSHTML.HTMLDocument
            doc.body.innerHTML = .responseText
        Else
            MsgBox "Error" & vbNewLine & "Ready state: " & .readyState & _
            vbNewLine & "HTTP request status: " & .Status
        End If
    End With
    
    Set post = doc
    Set doc = Nothing
    
End Function

Private Function nodeToDom(ByRef nodeStr As String) As MSHTML.HTMLDocument
    
    Dim doc As MSHTML.HTMLDocument
    
    Set doc = New MSHTML.HTMLDocument
    doc.body.innerHTML = nodeStr
    
    Set nodeToDom = doc
    Set doc = Nothing
    
End Function

Private Function domToList(ByRef query As String, ByRef doc As MSHTML.HTMLDocument, Optional ByVal elementAsArray As Boolean = False, Optional ByVal childQuery As String, Optional ByVal tabSep As Boolean = False) As Lists
    
    Dim j As Object
    Dim i
    Dim l As New Lists
    Dim l1 As New Lists
    
    Set j = doc.querySelectorAll(query)
    
    If elementAsArray Then
        If IsMissing(childQuery) Or childQuery = "" Then
            If tabSep Then
                For i = 0 To j.length - 1
                    l.add l1.fromArray(Split(j.Item(i).innerText, Chr(9)))
                Next i
            Else
                For i = 0 To j.length - 1
                    l.add l1.fromArray(Split(j.Item(i).innerText, Chr(13)))
                Next i
            End If
            
        Else
            For i = 0 To j.length - 1
                l.add domToList(childQuery, nodeToDom(j.Item(i).outerHTML))
            Next i
        End If
    Else
        For i = 0 To j.length - 1
            l.add j.Item(i).innerText
        Next i
    End If
    
    Set domToList = l
    Set j = Nothing
    Set l = Nothing
    Set l1 = Nothing
    
End Function