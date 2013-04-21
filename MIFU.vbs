mcloc=InputBox("Please enter you .minecraft location ex.) C:\Path\To\Your\multimcinstace")

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
	strFileURL = arrModList(0)
    For i = 1 to Ubound(arrModList)
		strHDLocation = mcloc & arrModList(i)
		download()
    Next
Loop

MsgBox("Finished Downloading Mods")
