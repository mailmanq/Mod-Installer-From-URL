Function download()

    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

    objXMLHTTP.open "GET", strFileURL, false
    objXMLHTTP.send()

    If objXMLHTTP.Status = 200 Then
      Set objADOStream = CreateObject("ADODB.Stream")
      objADOStream.Open
      objADOStream.Type = 1 

      objADOStream.Write objXMLHTTP.ResponseBody
      objADOStream.Position = 0    

      Set objFSO = Createobject("Scripting.FileSystemObject")
        If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
      Set objFSO = Nothing

      objADOStream.SaveToFile strHDLocation
      objADOStream.Close
      Set objADOStream = Nothing
    End if

    Set objXMLHTTP = Nothing
End Function

Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("Modlist.txt", ForReading)
Do Until objTextFile.AtEndOfStream
    strNextLine = objTextFile.Readline
    arrModList = Split(strNextLine , ",")
    Wscript.Echo "URL: " & arrModList(0)
    For i = 1 to Ubound(arrModList)
        Wscript.Echo "FILE: " & arrModList(i)
    Next
Loop

'Call download("http://files.minecraftforge.net/minecraftforge/minecraftforge-universal-latest.zip","minecraftforge-universal-latest.zip")
'mcloc=InputBox("Please enter you .minecraft location ex.) C:\Path\To\Your\multimcinstace")

'MCForge
'strFileURL = "http://files.minecraftforge.net/minecraftforge/minecraftforge-universal-latest.zip"
'strHDLocation = mcloc & "\instMods\" & "minecraftforge-universal-latest.zip"
'download()

'RedpowerCore
'strFileURL = "http://www.eloraam.com/files/43143756a7636620da44/RedPowerCore-2.0pr6.zip"
'strHDLocation = mcloc & "\minecraft\mods\" & "RedPowerCore-2.0pr6.zip"
'download()

'CodeChickenCore
'strFileURL = "http://www.chickenbones.craftsaddle.org/Files/goto.php?file=CodeChickenCore&version=1.5.1"
'strHDLocation = mcloc & "\minecraft\coremods\" & "CodeChickenCore.jar"
'download()

'NEI
'strFileURL = "http://www.chickenbones.craftsaddle.org/Files/goto.php?file=NotEnoughItems&version=1.5.1"
'strHDLocation = mcloc & "\minecraft\coremods\" & "NotEnoughItems.jar"
'download()

'MsgBox("Finished Downloading Mods")
