 function getCurrentRunFolder ()
	On Error  Goto 0 
  Dim objShell
    Set objShell = CreateObject("Wscript.Shell")

  Dim strPath
    strPath = Wscript.ScriptFullName
    ' Wscript.Echo  strPath  

  Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")

  Dim objFile 
  Set objFile = objFSO.GetFile(strPath)

  getCurrentRunFolder = objFSO.GetParentFolderName(objFile) 

  set objFile = Nothing
  set objFSO = Nothing 
  Set objShell = Nothing  
end function  

sub runScenario (vScrptFileName,vCurrPath)
 
		   Dim objFSO,objFile
       Dim vStrLine
       Dim  strCurCommands ,vCurrScrptFileName
       vCurrScrptFileName = vCurrPath & "\Commands\" & vScrptFileName 
        Wscript.Echo  vCurrScrptFileName 
       
        Set objFSO=CreateObject("Scripting.FileSystemObject")
            
		  if objFSO.FileExists(vCurrScrptFileName) then
        set objFile = objFSO.OpenTextFile(vCurrScrptFileName,1) 
           Dim objWshShell
           Set objWshShell = CreateObject("WScript.Shell") 
        Do Until objFile.AtEndOfStream
          vStrLine = objFile.ReadLine
          if instr(vStrLine,"#") = 0 then 
            strCurCommands = vCurrPath & "\svStressTool_run.vbs "   
            strCurCommands = "cscript " & strCurCommands  & " "  & vStrLine
            Wscript.Echo  strCurCommands
            call objWshShell.Run  (strCurCommands,2,TRUE )
          end if 
        Loop
        objFile.Close          
        set objFSO = Nothing 
        set objFile = Nothing  
        set objWshShell = Nothing 
       else 
          Wscript.Echo  "Can't find execute file  file " & vCurrScrptFileName
      end if
end sub 
 
 

Set args = Wscript.Arguments
For Each arg In args 
 ' writeConole  arg
  call runScenario (arg,getCurrentRunFolder())
Next


      


     


  



