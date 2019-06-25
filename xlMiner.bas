 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     Util Class xlMiner
'@author                                   Qiou Yang
'@lastUpdate                               25.06.2019
'                                          adapt to the new query api
'@TODO                                     add comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private http As Object
Private j As js

Enum fsType
    INCOME_STMT = 0 '"incomestatements"
    BALANCE_STMT = 1 ' "balancesheet"
    CASHFLOW_STMT = 2 ' "cashflow"
End Enum

Private Sub Class_Initialize()
    Set http = CreateObject("MSXML2.XMLHTTP.3.0")
    Set j = New js
    
    'so that not overwritten
    j.appendMode = False
End Sub

Private Sub Class_Terminate()
    Set http = Nothing
    Set j = Nothing
End Sub

Public Function profile(Optional ByVal code As String = "0") As Dicts
    
    Dim d As New Dicts
    Dim t As String
    Dim i
    
    code = Format(code, "000000")
    
    t = post("http://www.cninfo.com.cn/data/project/commonInterface", "mergerMark=sysapi1068&paramStr=scode=" & code, True)
    j.code = "function parser1(s){return parser(parserJSON(s)[0]);}"
    
    d.dict = j.js.Run("parser1", t)
    
    Set profile = d
    Set d = Nothing

End Function


Public Function fs(ByVal code As String, ByVal year As Integer, ByVal quarter As Integer, Optional ByVal mtype As Integer = fsType.INCOME_STMT) As Lists
    
    Dim l As New Lists
    Dim stype As String
    Dim i
    
    stype = Array("sysapi1075", "sysapi1077", "sysapi1076")(mtype)
    code = Format(code, "000000")
    
    Dim t As String
    t = post("http://www.cninfo.com.cn/data/project/commonInterface", "mergerMark=" & stype & "&paramStr=scode=" & code & ";rtype=" & quarter & ";sign=1", True)
    
    
    j.code = "function parser1(s){return parser(parserJSON(s));}"
    
    For Each i In j.js.Run("parser1", t)
        Dim tmp As New Lists
        l.add tmp.fromArray(Array(i("index"), i("" & year)))
        Set tmp = Nothing
    Next i

    Set fs = l
    Set l = Nothing
   
    
End Function

'@return Array of 3 elements, with the first element of desc and second list of content, third next fullurl
Public Function DE_law(Optional ByVal law As String = "hgb", Optional ByVal parag As String = "1", Optional ByVal tillEnd As Boolean = False, Optional ByVal fullUrl As String = "") As Variant
    On Error GoTo hdl  ' wenn kein Weiter vorhanden
    
        Dim root As String
        root = "https://www.gesetze-im-internet.de/" & StrConv(law, vbLowerCase) & "/"
        
        Dim doc As MSHTML.HTMLDocument
        
        If fullUrl = "" Then
            fullUrl = root & "__" & parag & ".html"
        End If
        
        Set doc = post(fullUrl)

        Dim title As String
        title = doc.querySelector(".jnentitel").innerText
        
        Dim l As New Lists
        Dim i, v
        
        Dim j As Object
        Set j = doc.querySelectorAll(".jurAbsatz")  ' for each not fully supported by querySelectorAll
        
        For i = 0 To j.length - 1
             l.add j.Item(i).innerText
        Next i
        
        Dim u As String
        u = doc.querySelector("#blaettern_weiter > a").getAttribute("href")
        
        If InStr(u, "about:") Then
            u = Right(u, Len(u) - 6)
        End If
           
        If Not tillEnd Then
            DE_law = Array(parag & "-" & title, l, root & u)
        Else
            Dim l1 As New Lists
            Dim this
            
            this = Array(parag & "-" & title, l, root & u)
            l1.add this
            
            Do While True
                this = DE_law(fullUrl:=this(2))
                l1.add this
            Loop
    
hdl:
            Set DE_law = l1
            Set l1 = Nothing
           
        End If
        
        Set l = Nothing
        Set j = Nothing
        Set doc = Nothing

End Function

'  https://analystcave.com/vba-reference-functions/vba-string-functions/vba-strconv-function/
' locale code of PRC 804
Private Function recode(src As String, Optional fromCharSetLocale As Long = &H804, Optional toCharSetLocale As Long = &H403) As String

    recode = StrConv(StrConv(src, vbFromUnicode, fromCharSetLocale), vbUnicode, toCharSetLocale)

End Function

Private Function post(ByVal url As String, Optional ByVal data As String, Optional ByVal asText As Boolean = False, Optional ByVal verb As String = "POST")
    With http
        .Open verb, url, False
        .setRequestHeader "content-type", "application/x-www-form-urlencoded; charset=utf-8"
        If IsMissing(data) Or data = "" Then
            .send
        Else
            .send data
        End If
        
        If .readyState = 4 And .Status = 200 Then
            If asText Then
                post = Trim(.responseText)
            Else
                Dim doc As MSHTML.HTMLDocument
                Set doc = New MSHTML.HTMLDocument
                
                
                doc.body.innerHTML = Trim(.responseText)
                
                Set post = doc
                Set doc = Nothing
            End If
        Else
            MsgBox "Error" & vbNewLine & "Ready state: " & .readyState & _
            vbNewLine & "HTTP request status: " & .Status
        End If
    End With
    
End Function

Private Function nodeToDom(ByRef nodeStr As String) As MSHTML.HTMLDocument
    
    Dim doc As MSHTML.HTMLDocument
    
    Set doc = New MSHTML.HTMLDocument
    doc.body.innerHTML = nodeStr
    
    Set nodeToDom = doc
    Set doc = Nothing
    
End Function

Private Function domToList(ByRef Query As String, ByRef doc As MSHTML.HTMLDocument, Optional ByVal elementAsArray As Boolean = False, Optional ByVal childQuery As String, Optional ByVal tabSep As Boolean = False) As Lists
    
    Dim j As Object
    Dim i
    Dim l As New Lists
    Dim l1 As New Lists
    
    Set j = doc.querySelectorAll(Query)
    
    If j.length > 0 Then
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
    End If
    
    Set domToList = l
    Set j = Nothing
    Set l = Nothing
    Set l1 = Nothing
    
End Function
