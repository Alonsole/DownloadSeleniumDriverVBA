Attribute VB_Name = "Module1"
Sub errordownloaddriver()
Dim driver As New Selenium.ChromeDriver
Dim request As Object, file As Object
Dim response As String, desktopPath As String
Dim localZipFile As Variant, destFolder As Variant
Dim Sh As Object, ZipNamespace As Object, Item As Object
'Задать маршрут (где расположен Selenium)
selenium_path = "C:\Users\Maksim\AppData\Local\SeleniumBasic\"
'Настройка подключения к браузеру
driver.SetCapability "debuggerAddress", "localhost:9222"

'Обработка ошибки
On Error Resume Next
    driver.Start ("CHROME")
If err.Number = 33 Then
    MsgBox "Ошибка при создании сессии. Будет удален старый драйвер, скачан новый драйвер. Можете повторить запрос по выполнению операции", vbExclamation
     
    On Error GoTo err
    'Блок загрузки драйвера
        Url = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_STABLE"
    ' Отправить GET-запрос
    Set request = CreateObject("MSXML2.XMLHTTP")
        request.Open "GET", Url, False
        request.Send
        response = request.ResponseText
        'Debug.Print response
        download = "https://storage.googleapis.com/chrome-for-testing-public/" & response & "/win64/chromedriver-win64.zip"
        request.Open "GET", download, False
        request.Send
        'Сохранить файл
    Set file = CreateObject("ADODB.Stream")
        file.Type = 1
        file.Open
        file.Write request.ResponseBody
        Name = response & ".zip"
        file.SaveToFile selenium_path & Name, 2
        request.Abort
        file.Close
        Set request = Nothing
        Set file = Nothing
        
    'Блок удаления предыдущего драйвера
    Set wsh = CreateObject("WScript.Shell")
        wsh.Run "powershell.exe Stop-Process -name chromedriver -force", vbHide
        Application.Wait (Now + TimeValue("0:00:03"))
        Kill selenium_path & "chromedriver.exe"
        
    'Блок распаковки
        localZipFile = selenium_path & Name
    Set Sh = CreateObject("Shell.Application")
    Set ZipNamespace = Sh.Namespace(localZipFile)
        Sh.Namespace(selenium_path).CopyHere ZipNamespace.Items.Item.Path & "\chromedriver-win64\chromedriver.exe"
        
    'Блок удаления архива с драйвером
        Kill selenium_path & Name
    'Информирование
        MsgBox "Операция по загрузке и замене Драйвера выполнена", vbInformation
End If

'Блок запуска после устанения ошибки
On Error GoTo 0
    driver.Start ("CHROME")
    driver.Get ("https://www.ya.ru/")
    MsgBox "Успешный запуск", vbInformation
Exit Sub
err:
    MsgBox "Операция завершена с ошибкой", vbCritical
End Sub
