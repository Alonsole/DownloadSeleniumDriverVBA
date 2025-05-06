Attribute VB_Name = "Module1"
Sub errordownloaddriver()
Dim driver As New Selenium.ChromeDriver
Dim request As Object, file As Object
Dim response As String, desktopPath As String
Dim localZipFile As Variant, destFolder As Variant
Dim Sh As Object, ZipNamespace As Object, Item As Object
'������ ������� (��� ���������� Selenium)
selenium_path = "C:\Users\Maksim\AppData\Local\SeleniumBasic\"
'��������� ����������� � ��������
driver.SetCapability "debuggerAddress", "localhost:9222"

'��������� ������
On Error Resume Next
    driver.Start ("CHROME")
If err.Number = 33 Then
    MsgBox "������ ��� �������� ������. ����� ������ ������ �������, ������ ����� �������. ������ ��������� ������ �� ���������� ��������", vbExclamation
    
    'On Error GoTo err
    '���� �������� ��������
        Url = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_STABLE"
    ' ��������� GET-������
    Set request = CreateObject("MSXML2.XMLHTTP")
        request.Open "GET", Url, False
        request.Send
        response = request.ResponseText
        'Debug.Print response
        download = "https://storage.googleapis.com/chrome-for-testing-public/" & response & "/win64/chromedriver-win64.zip"
        request.Open "GET", download, False
        request.Send
        '��������� ����
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
        
    '���� �������� ����������� ��������
    Set wsh = CreateObject("WScript.Shell")
        wsh.Run "powershell.exe Stop-Process -name chromedriver -force", vbHide
        Application.Wait (Now + TimeValue("0:00:03"))
        Kill selenium_path & "chromedriver.exe"
        
    '���� ����������
        localZipFile = selenium_path & Name
    Set Sh = CreateObject("Shell.Application")
    Set ZipNamespace = Sh.Namespace(localZipFile)
        Sh.Namespace(selenium_path).CopyHere ZipNamespace.Items.Item.Path & "\chromedriver-win64\chromedriver.exe"
        
    '���� �������� ������ � ���������
        Kill selenium_path & Name
    '��������������
        MsgBox "�������� �� �������� � ������ �������� ���������", vbInformation
End If

'���� ������� ����� ��������� ������
On Error GoTo 0
    driver.Start ("CHROME")
    driver.Get ("https://www.ya.ru/")
    MsgBox "�������� ������", vbInformation
Exit Sub
err:
    MsgBox "�������� ��������� � �������", vbCritical
End Sub
