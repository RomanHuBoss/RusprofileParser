' ///////////////////////////////////////////////
' R.Rabinovich 2019 (roman@web-line.ru)
'
' скрипт на вход ожидает следующие параметры:
' -i=[Integer] - ИНН (целое число)
' -o=[Integer] - ОГРН (целое число)
' -s=[SiteName] - название сайта (опционально), с которого будет дёргать данные. По умолчанию - rusprofile
' -f=[FolderPath] - путь к папке, куда будет сохраняться файл (по умолчанию - в ту же, где и скрипт)
' //////////////////////////////////////////////

' глобальные переменные
Dim argumentsParseError                                      ' текст ошибки парсинга запроса
Dim requestedInn                                             ' запрашиваемый ИНН
Dim requestedOgrn                                            ' запрашиваемый ОГРН
Dim requestedSiteName : requestedSiteName = "rusprofile"     ' имя запрашиваемого сайта
Dim destinationFolder : destinationFolder = ""               ' папка, в которую будет сохраняться результат

' ассоциативный массив, связывающий имена сайтов с "шаблонами" URL,
' в которые при обработке вместо [INN] будут подставлены ИНН, вместо [OGRN] ОГРН, вместо [ANY] любой из этих параметров
SET sites = CreateObject("Scripting.Dictionary")
sites.Add "rusprofile", "https://www.rusprofile.ru/search?query=[ANY]"

' ассоциативный массив, заполняемый в ходе парсинга данными организации
SET orgData = CreateObject("Scripting.Dictionary")

If CheckArguments = True Then
   HandleRequest()
Else
   WScript.Echo argumentsParseError
End If


' ф-ция запрашивает данные ИНН на заданном сайте
Function HandleRequest()
    Dim url

    url = Replace(sites.Item(requestedSiteName), "[INN]", requestedInn)
    url = Replace(url, "[OGRN]", requestedOgrn)
    if (requestedInn <> "") Then
        url = Replace(url, "[ANY]", requestedInn)
    Else
        url = Replace(url, "[ANY]", requestedOgrn)
    End If

    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    oXMLHTTP.Open "GET", url, False
    oXMLHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:62.0) Gecko/20100101 Firefox/62.0"
    oXMLHTTP.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    oXMLHTTP.setRequestHeader "Accept-Language", "ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3"
    oXMLHTTP.setRequestHeader "Accept-Encoding", "gzip, deflate, br"
    oXMLHTTP.setRequestHeader "Connection", "keep-alive"
    oXMLHTTP.Send

    If oXMLHTTP.Status <> 200 Then
        WScript.Echo "Can't fetch data from " + url
        Exit Function
    End if

    Set html = CreateObject("htmlfile")
    html.write oXMLHTTP.responseText

    ' для сайта rusprofile.ru перезагрузим страницу в режиме "для печати"
    If requestedSiteName = "rusprofile" Then
        ' вот тут перебираем все элементы с тегом линк на первоначально загруженной странице
        For Each item In html.getElementsByTagName("link")
            ' находим такой, у которого есть свойство canonical
            If item.getAttribute("rel") = "canonical" Then
                ' у него и тащим канонический URL, к которому дописываем опцию print
                url = item.getAttribute("href") + "?print=1"
                oXMLHTTP.Open "GET", url, False
                oXMLHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:62.0) Gecko/20100101 Firefox/62.0"
                oXMLHTTP.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                oXMLHTTP.setRequestHeader "Accept-Language", "ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3"
                oXMLHTTP.setRequestHeader "Accept-Encoding", "gzip, deflate, br"
                oXMLHTTP.setRequestHeader "Connection", "keep-alive"
                oXMLHTTP.Send

                If oXMLHTTP.Status <> 200 Then
                    WScript.Echo "Can't fetch data from " + url
                    Exit Function
                End If

                html.write oXMLHTTP.responseText

                Exit For
            End If
        Next
    End If

    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write oXMLHTTP.responseBody

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' нет папки, в которую будем писать, тогда создадим
    If len(destinationFolder) <> 0 AND fso.FolderExists(destinationFolder) = False Then
         fso.CreateFolder(destinationFolder)
    End If

    ' пути к 2 сохраняемым файлам: оригинал в html, результат парсинга - в txt
    Dim htmlFilePath, txtFilePath
    If requestedInn <> "" Then
        htmlFilePath = destinationFolder + "/" + requestedInn + ".html"
        txtFilePath = destinationFolder + "/" + requestedInn + ".txt"
    Else
        htmlFilePath = destinationFolder + "/" + requestedOgrn + ".html"
        txtFilePath = destinationFolder + "/" + requestedOgrn + ".txt"
    End If

    ' бьем старые файлы, если имеются
    If fso.FileExists(htmlFilePath) = True Then
        fso.DeleteFile(htmlFilePath)
    End If

    ' бьем старые файлы, если имеются
    If fso.FileExists(txtFilePath) = True Then
        fso.DeleteFile(txtFilePath)
    End If

    ' для будущих поколений сохраним оригинальный HTML, пришедший с сервера
    oStream.SaveToFile htmlFilePath
    oStream.Close

    ' начинаем парсить HTML-файл в зависимости от того, с каким сайтом работаем
    Set html = CreateObject("htmlfile")
    html.write oXMLHTTP.responseText
    ParseHtmlPage(html)

    ' собственно сваливаем все, что насобирали, в файл
    Dim textToFile
    For Each key In orgData
        textToFile = textToFile + (key + "=" + orgData.Item(key) & vbCrLf)
    Next

    Set txtFile = fso.CreateTextFile(txtFilePath, True)
    txtFile.Write textToFile
    txtFile.Close


End Function

' ф-ция парсит HTML-файл, полученный с сайта
Function ParseHtmlPage(html)
    Dim cellName, cellValue

    ' бежим по контенту сайта rusprofile
    If requestedSiteName = "rusprofile" Then

        For Each tag In html.getElementById("anketa").All
            ' краткое наименование организации
            If lcase(tag.tagName) = "h1" Then
                orgData.Add "orgShortName", tag.innerText
            End If

            ' полное наименование организации
            If lcase(tag.tagName) = "h2" AND tag.innerText <> ConvertUtfTo1251("ЛИКВИДАЦИЯ") Then
                orgData.Add "orgFullName", tag.innerText
            ElseIf lcase(tag.tagName) = "h2" AND tag.innerText = ConvertUtfTo1251("ЛИКВИДАЦИЯ") Then
                orgData.Add "orgStatus", ConvertUtfTo1251("Ликвидация")
            End If

            ' адрес организации
            If lcase(tag.tagName) = "p" AND tag.className = "darkline" Then
                orgData.Add "orgAddress", tag.innerText
            End If

            ' руководитель организации
            If lcase(tag.tagName) = "p" AND tag.className = "lightline" Then
                orgData.Add "orgBoss", tag.innerText
            End If

            ' на сайте дважды используется тег id="requisites" в этой секции и в самом BODY. приходится извращаться
            If tag.getAttribute("id") = "requisites" Then
                For Each item In tag.All
                    If lcase(item.tagName) = "div" AND item.className = "colunit" Then
                        For Each subItem In item.getElementsByTagName("TR")
                            If subItem.getElementsByTagName("TD").length = 2 Then
                                cellName = subItem.getElementsByTagName("TD")(0).innerText
                                cellValue = subItem.getElementsByTagName("TD")(1).innerText

                                ' ФИО
                                If cellName = ConvertUtfTo1251("ФИО") Then
                                    orgData.Add "FIO", cellValue
                                End If

                                ' Регион
                                If cellName = ConvertUtfTo1251("Регион") Then
                                    orgData.Add "Region", cellValue
                                End If

                                ' Адрес
                                If cellName = ConvertUtfTo1251("Адрес") Then
                                    orgData.Add "Address", cellValue
                                End If

                                ' Пол
                                If cellName = ConvertUtfTo1251("Пол") Then
                                    orgData.Add "Sex", cellValue
                                End If

                                ' Гражданство
                                If cellName = ConvertUtfTo1251("Гражданство") Then
                                    orgData.Add "Citizenship", cellValue
                                End If

                            End If
                        Next
                    End If
                Next
            End If
        Next

        For Each tag In html.All
            If tag.getAttribute("id") = "requisites" Then
                For Each item In tag.All
                    If lcase(item.tagName) = "div" AND item.className = "colunit" Then
                        If item.getElementsByTagName("H2").length = 1 Then

                            Dim sectionTitle : sectionTitle = item.getElementsByTagName("H2")(0).innerText
                            Dim prefix

                            ' там множество подсекций со сходными реквизитами, поэтому каждой дадим по префиксу
                            If sectionTitle = ConvertUtfTo1251("РЕКВИЗИТЫ ОРГАНИЗАЦИИ") Then
                                orgData.Add "type", ConvertUtfTo1251("Организация")
                                prefix = "common"
                            Elseif sectionTitle = ConvertUtfTo1251("РЕКВИЗИТЫ ИП") Then
                                orgData.Add "type", ConvertUtfTo1251("Индивидуальный предприниматель")
                                prefix = "common"
                            Elseif sectionTitle = ConvertUtfTo1251("СВЕДЕНИЯ РОССТАТА") Then
                                prefix = "rosstat"
                            Elseif sectionTitle = ConvertUtfTo1251("РЕГИСТРАЦИЯ В ФНС") Then
                                prefix = "fnsreg"
                            Elseif sectionTitle = ConvertUtfTo1251("РЕГИСТРАЦИЯ В ПФР") Then
                                prefix = "pfrreg"
                            Elseif sectionTitle = ConvertUtfTo1251("РЕГИСТРАЦИЯ В ФСС") Then
                                prefix = "fssreg"
                            Elseif sectionTitle = ConvertUtfTo1251("СВЕДЕНИЯ РЕЕСТРА МСП") Then
                                prefix = "mspreg"
                            Else
                                prefix = "unknown"
                            End If

                            For Each subItem In item.getElementsByTagName("TR")
                                If subItem.getElementsByTagName("TD").length = 2 Then
                                    cellName = subItem.getElementsByTagName("TD")(0).innerText
                                    cellValue = subItem.getElementsByTagName("TD")(1).innerText

                                    ' ОГРН
                                    If cellName = ConvertUtfTo1251("ОГРН") Then
                                        orgData.Add prefix + "OGRN", cellValue
                                    End If

                                    ' ОГРНИП
                                    If cellName = ConvertUtfTo1251("ОГРНИП") Then
                                        orgData.Add prefix + "OGRNIP", cellValue
                                    End If

                                    ' Вид предпринимательства
                                    If cellName = ConvertUtfTo1251("Вид предпринимательства") Then
                                        orgData.Add prefix + "BusinessKind", cellValue
                                    End If

                                    ' ИНН
                                    If cellName = ConvertUtfTo1251("ИНН") Then
                                        orgData.Add prefix + "INN", cellValue
                                    End If

                                    ' КПП
                                    If cellName = ConvertUtfTo1251("КПП") Then
                                        orgData.Add prefix + "KPP", cellValue
                                    End If

                                    ' Дата постановки на учёт
                                    If cellName = ConvertUtfTo1251("Дата постановки на учёт") Then
                                        orgData.Add prefix + "UchetDate", cellValue
                                    End If

                                    ' Налоговый орган
                                    If cellName = ConvertUtfTo1251("Налоговый орган") Then
                                        orgData.Add prefix + "TaxBranch", cellValue
                                    End If

                                    ' Налоговый орган
                                    If cellName = ConvertUtfTo1251("Наименование налогового органа") Then
                                        orgData.Add prefix + "TaxBranchName", cellValue
                                    End If

                                    ' Уставный капитал
                                    If cellName = ConvertUtfTo1251("Уставный капитал") Then
                                        orgData.Add prefix + "BaseFunds", cellValue
                                    End If

                                    ' ОКПО
                                    If cellName = ConvertUtfTo1251("ОКПО") Then
                                        orgData.Add prefix + "OKPO", cellValue
                                    End If

                                    ' ОКАТО
                                    If cellName = ConvertUtfTo1251("ОКАТО") Then
                                        orgData.Add prefix + "OKATO", cellValue
                                    End If

                                    ' ОКОГУ
                                    If cellName = ConvertUtfTo1251("ОКОГУ") Then
                                        orgData.Add prefix + "OKOGU", cellValue
                                    End If

                                    ' ОКТМО
                                    If cellName = ConvertUtfTo1251("ОКТМО") Then
                                        orgData.Add prefix + "OKTMO", cellValue
                                    End If

                                    ' ОКФС
                                    If cellName = ConvertUtfTo1251("ОКФС") Then
                                        orgData.Add prefix + "OKFS", cellValue
                                    End If

                                    ' Дата регистрации
                                    If cellName = ConvertUtfTo1251("Дата регистрации") Then
                                        orgData.Add prefix + "RegDate", cellValue
                                    End If

                                    ' Регистратор
                                    If cellName = ConvertUtfTo1251("Регистратор") Then
                                        orgData.Add prefix + "Registrator", cellValue
                                    End If

                                    ' Адрес регистратора
                                    If cellName = ConvertUtfTo1251("Адрес регистратора") Then
                                        orgData.Add prefix + "RegistratorAddress", cellValue
                                    End If

                                    ' Регистрационный номер
                                    If cellName = ConvertUtfTo1251("Регистрационный номер") Then
                                        orgData.Add prefix + "RegNumber", cellValue
                                    End If

                                    ' Наименование территориального органа
                                    If cellName = ConvertUtfTo1251("Наименование территориального органа") Then
                                        orgData.Add prefix + "TerritorialBranch", cellValue
                                    End If

                                    ' Дата включения в МСП
                                    If cellName = ConvertUtfTo1251("Дата включения в МСП") Then
                                        orgData.Add prefix + "MspIncludeDate", cellValue
                                    End If

                                    ' Категория субъекта МСП
                                    If cellName = ConvertUtfTo1251("Категория субъекта МСП") Then
                                        orgData.Add prefix + "MspSubjectCategory", cellValue
                                    End If

                                End If
                            Next
                        End If
                    End If
                Next
            End If
        Next
    End If

End Function


' ф-ция проверяет наличие и корректность параметров запроса
Function CheckArguments()
    Dim argsCount : argsCount = WScript.Arguments.count

    If argsCount < 1 Then
        CheckArguments = False
        argumentsParseError = "Minimum arguments number should be more or equal to 1"
    Else
        For Each arg In WScript.Arguments
            If Left(arg, 3) = "-i=" Then
                requestedInn = Mid(arg, 4)
            End If

            If Left(arg, 3) = "-o=" Then
                requestedOgrn = Mid(arg, 4)
            End If

            If Left(arg, 3) = "-s=" Then
                requestedSiteName = Mid(arg, 4)
            End If

            If Left(arg, 3) = "-f=" Then
                destinationFolder = Mid(arg, 4)
            End If
        Next

        If requestedInn = "" AND requestedOgrn = "" Then
            argumentsParseError = "Neither INN, nor OGRN were requested"
            CheckArguments = False
        ElseIf requestedInn <> "" AND CheckInnCorrectness(requestedInn) = False Then
            argumentsParseError = "Wrong INN value requested"
            CheckArguments = False
        ElseIf requestedOgrn <> "" AND CheckOgrnCorrectness(requestedOgrn) = False Then
            argumentsParseError = "Wrong OGRN value requested"
            CheckArguments = False
        ElseIf CheckSiteCorrectness(requestedSiteName) = False Then
            argumentsParseError = "Wrong site name requested"
            CheckArguments = False
        ElseIf CheckDestinationFolder(destinationFolder) = False Then
            argumentsParseError = "Wrong destination folder"
            CheckArguments = False
        Else
            CheckArguments = True
        End If

    End If
End Function

' ф-ция проверяет корректность ИНН
Function CheckInnCorrectness(inn)
    If IsNumeric(inn) Then
        If inn > 0 Then
            CheckInnCorrectness = True
        Else
            CheckInnCorrectness = False
        End If
    Else
        CheckInnCorrectness = False
    End If
End Function

' ф-ция проверяет корректность ОГРН
Function CheckOgrnCorrectness(ogrn)
    If IsNumeric(ogrn) Then
        If ogrn > 0 Then
            CheckOgrnCorrectness = True
        Else
            CheckOgrnCorrectness = False
        End If
    Else
        CheckOgrnCorrectness = False
    End If
End Function


' ф-ция проверяет корректность имени запрашиваемого сайта
Function CheckSiteCorrectness(site)
    CheckSiteCorrectness = sites.Exists(site)
End Function

' ф-ция проверяет наличие папки назначения и пытается ее создать при отсутствии
Function CheckDestinationFolder(destinationFolder)

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' нет папки, в которую будем писать, тогда создадим
    If len(destinationFolder) <> 0 AND fso.FolderExists(destinationFolder) = False Then
        fso.CreateFolder(destinationFolder)

        ' папка не создана
        If fso.FolderExists(destinationFolder) = False Then
            CheckDestinationFolder = False
            Exit Function
        End If
    End If

    CheckDestinationFolder = True

End Function

' перекодировка строки из CP-1251 в UTF-8
Function Convert1251ToUtf(sIn)
    Set Recode = CreateObject("ADODB.Stream")
    Recode.Open
    Recode.CharSet = "UTF-8"
    Recode.WriteText sIn
    Recode.Position = 0
    Recode.CharSet = "windows-1251"
    Convert1251ToUtf = Recode.ReadText
    Recode.Close
End Function

' перекодировка строки из UTF-8 в CP-1251
Function ConvertUtfTo1251(sIn)
    Set Recode = CreateObject("ADODB.Stream")
    Recode.Open
    Recode.CharSet = "windows-1251"
    Recode.WriteText sIn
    Recode.Position = 0
    Recode.CharSet = "UTF-8"
    ConvertUtfTo1251 = Recode.ReadText
    Recode.Close
End Function
