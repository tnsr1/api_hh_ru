'autor alex; alexeyzapolskiy@gmail.com
'24.06.2024
Sub vvv()
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    timeout = 2000 'milliseconds
    http.setTimeouts timeout, timeout, timeout, timeout
'    http.Option(2) = 0 'Замена дефолтного значения 65001 для отмены преобразования в utf-8
    Dim url0 As String, url_ As String
    url0 = "https://api.hh.ru/vacancies?text=NAME:(Программист) and DESCRIPTION:(NOT intermediate)&area=1&only_with_salary=true&no_magic=true&salary=100000&currency_code=RUR&period=30&label=not_from_agency&order_by=publication_time"
    url0 = URLEncode(url0)
    http.Open "get", url0
    http.send
    Text = http.responseText
    If InStr(Text, "errors") > 0 Then
        Debug.Print Text
        Stop
    Else
        If Text <> "" Then
            Set qwe = JsonConverter.ParseJson(Text)
        End If
    End If
    CountV = qwe("found")
    CountP = qwe("pages")
    isk = 1
On Error GoTo AfterSk
    ThisWorkbook.Worksheets(2).Range("B:B").Font.ColorIndex = 0
    Dim wsMySkills As Worksheet
    Set wsMySkills = ThisWorkbook.Worksheets("Мои навыки")
    For pg = 1 To CountP
        If pg > 1 Then
            url_ = url0 & "&page=" & pg
            http.Open "get", url_
            http.send
            Text = http.responseText
            Set qwe = JsonConverter.ParseJson(Text)
        End If
        For i = 1 To 20
            ii = (pg - 1) * 20 + i
            Set Item = qwe("items")(i)
            url1 = Item("alternate_url")
            ThisWorkbook.Worksheets(2).Cells(ii + isk, 1) = Item("name")
            ThisWorkbook.Worksheets(2).Cells(ii + isk, 3) = url1
            ThisWorkbook.Worksheets(2).Cells(ii + isk, 1).Font.Bold = True
            ThisWorkbook.Worksheets(2).Cells(ii + isk, 1).Font.Size = 14
            ThisWorkbook.Worksheets(2).Cells(ii + isk, 3).Font.Bold = True
            url_ = Item("url")
            url_ = Replace(url_, "?host=hh.ru", "")
            http.Open "get", url_
            http.send
            Text = http.responseText
            Set vak = JsonConverter.ParseJson(Text)
            Set keySkills = vak("key_skills")
            CountSk = keySkills.Count
            If CountSk > 0 Then
                For jj = 1 To CountSk
                    If jj <> 1 Then isk = isk + 1
                    ThisWorkbook.Worksheets(2).Cells(ii + isk, 1) = Item("name")
                    keySkill1 = keySkills(jj)("name")
                    ThisWorkbook.Worksheets(2).Cells(ii + isk, 2) = keySkill1
                    ThisWorkbook.Worksheets(2).Cells(ii + isk, 2).Font.Italic = True
                    ThisWorkbook.Worksheets(2).Cells(ii + isk, 3) = url1
                    wsMySkills.Cells(1, 2).FormulaR1C1 = "=VLOOKUP(""" & keySkill1 & """,R1C1:R100C1,1,FALSE)"
                    If keySkill1 = wsMySkills.Cells(1, 2).Text Then
                        ThisWorkbook.Worksheets(2).Cells(ii + isk, 2).Font.ColorIndex = 10
                        ThisWorkbook.Worksheets(2).Cells(ii + isk, 4).Value = 1
                    Else
                        ThisWorkbook.Worksheets(2).Cells(ii + isk, 2).Font.ColorIndex = 3
                        ThisWorkbook.Worksheets(2).Cells(ii + isk, 4).Value = 0
                    End If
                Next jj
'            Else
'                ThisWorkbook.Worksheets(2).Cells(2 + (ii - 1) * 3, 1) = vak("description")
'                ThisWorkbook.Worksheets(2).Cells(2 + (ii - 1) * 3, 1).Select
'                Rows("2 + (ii - 1) * 3:2 + (ii - 1) * 3").EntireRow.AutoFit
            End If
AfterSk:
        If Err.Number <> 0 Then
            Debug.Print Err.Description
            'Stop
            Resume Next
            Err.Clear
        End If
            DoEvents
        Next i
    Next pg
    Stop
End Sub

Public Function URLEncode(ByRef txt As String) As String
    Dim buffer As String, i As Long, c As Long, n As Long
    buffer = String$(Len(txt) * 12, "%")
 
    For i = 1 To Len(txt)
           c = AscW(Mid$(txt, i, 1)) And 65535
    
           Select Case c
               'Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
               Case 38, 40, 41, 47 To 58, 61, 63, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
                   n = n + 1
                   Mid$(buffer, n) = ChrW(c)
               Case Is <= 127            ' Escaped UTF-8 1 bytes U+0000 to U+007F '
                   n = n + 3
                   Mid$(buffer, n - 1) = Right$(Hex$(256 + c), 2)
               Case Is <= 2047           ' Escaped UTF-8 2 bytes U+0080 to U+07FF '
                   n = n + 6
                   Mid$(buffer, n - 4) = Hex$(192 + (c \ 64))
                   Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
               Case 55296 To 57343       ' Escaped UTF-8 4 bytes U+010000 to U+10FFFF '
                   i = i + 1
                   c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, i, 1)) And 1023)
                   n = n + 12
                   Mid$(buffer, n - 10) = Hex$(240 + (c \ 262144))
                   Mid$(buffer, n - 7) = Hex$(128 + ((c \ 4096) Mod 64))
                   Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                   Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
               Case Else                 ' Escaped UTF-8 3 bytes U+0800 to U+FFFF '
                   n = n + 9
                   Mid$(buffer, n - 7) = Hex$(224 + (c \ 4096))
                   Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                   Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
           End Select
    Next
    URLEncode = Left$(buffer, n)
End Function
