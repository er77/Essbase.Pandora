
Sub includeFile(fCurrFile)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fCurrFile).readAll()
    End With
End Sub

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
 

sub writeConole  (LogMessage )
  Wscript.Echo Chr(13) & getRightTime() & ";" & LogMessage
end sub

sub executeCommnad (vCurrCommand,strCurrPath)      
		   Dim arrParametres,strCurCommands
       arrParametres = split (vCurrCommand,";") 
        '#ScenarioName;times;mode;delay
        'RU3ACTIV.scn01;100;sunc;2

       if ubound (arrParametres) > 2  then 
         strCurCommands = strCurrPath & "\svStressTool_main.vbs " &  arrParametres(0)  
         strCurCommands = "cscript " & strCurCommands  & " "  
         Dim strLogFile
         strLogFile =  strCurrPath & "\logs\" & vLogPrefix & replace(arrParametres(0),".","_") 
         strCurCommands =  strCurCommands & strLogFile & "_all.log" 
         writeConole strCurCommands           
       Dim objWshShell
      Set objWshShell = CreateObject("WScript.Shell")         
         for i = 1 to arrParametres(1)              
           if instr(ucase(arrParametres(2)),"ASYNC") then                          
                  call objWshShell.Run ( strCurCommands,2,FALSE )
           else             
                  call objWshShell.Run  (strCurCommands,2,TRUE )
           end if            
           WScript.Sleep arrParametres(3) * 1000
         next 
       end if  

      set objWshShell = nothing 
         
end sub 

dim strPath
 
 strPath =  getCurrentRunFolder()

 includeFile strPath & "\svStressTool_svc_log.vbs"

dim  vLogPrefix
 vLogPrefix = getRightTime()
 vLogPrefix = replace(vLogPrefix," ","")
 vLogPrefix = replace(vLogPrefix,":","")

Set args = Wscript.Arguments
For Each arg In args 
 ' writeConole  arg
  call executeCommnad (arg,strPath)
Next


      


     


  



