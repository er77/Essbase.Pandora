Sub includeFile(fCurrFile)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fCurrFile).readAll()
    End With
End Sub

Dim  strLogFileName

sub makeLogs (vCurrLogMessage) 
  if (instr (vCurrLogMessage,"SLEEP") = 0 ) then
   if (instr ( ucase(strLogFileName) , "CONN") > 0 ) then 
      call writeConole (vCurrLogMessage )
    else 
      call WriteFileLog (strLogFileName,vCurrLogMessage)
    end if  
  end if  
end sub 


Dim vSID,vSSOCore,vSSOApp,vSSODB
Dim vAPS, vESB
Dim vLogin, vPass
Dim vApp, vDB


sub makeLogin 
    makeLogs "Start Login "      
       vSID     =""
       vSSOCore ="" 
       vSSOApp  ="" 
       vSSODB   =""

        vSid = getSID( vAPS ,vLogin,vPass)      
        vSSOCore = getSSO(vAPS,vSid)     
        vSSOApp  = getOpennedApplication(vAPS,vSID,vSSOCore ,vESB,vApp)	  
        vSSODB   = getOpennedCube(vAPS,vSID,vSSOCore ,vESB,vApp,vDB) 	
            
     makeLogs "Stop Login "  
end sub 

sub doLogin (vCurrLoginFileName)
 
		   Dim objFSO,objFile
        Set objFSO=CreateObject("Scripting.FileSystemObject")
        Dim vStrLine     
        Dim vConnCount 
		  if objFSO.FileExists(vCurrLoginFileName) then
        set objFile = objFSO.OpenTextFile(vCurrLoginFileName,1)
'APS=http://wedcb787.frmon.danet:13080/aps/SmartView
'ESB=wedcb591.frmon.danet:1424
'LOG=rurasyukev
'PAS=password 
'APP=RU3ACTIV
'DBS=RU3ACTIV    
        vConnCount = 0 
        Do Until objFile.AtEndOfStream
          vStrLine = objFile.ReadLine
         ' makeLogs vStrLine
          vCurrComand = split(vStrLine,"=")          
          if Ubound (vCurrComand)=1 then        
            Select Case ucase(vCurrComand(0))
               Case "APS"
                    vAPS   = vCurrComand(1)  
                    vConnCount = vConnCount + 1 
               Case "ESB"
                    vESB   = vCurrComand(1)     
                    vConnCount = vConnCount + 1               
               Case "LOG"
                    vLogin = vCurrComand(1)  
                    vConnCount = vConnCount + 1 
               Case "PAS"
                    vPass  = vCurrComand(1)     
                    vConnCount = vConnCount + 1 
               Case "APP"
                    vApp   = vCurrComand(1)  
                    vConnCount = vConnCount + 1 
               Case "DBS"
                    vDB    = vCurrComand(1)     
                    vConnCount = vConnCount + 1               
               Case Else
                     makeLogs "error " &  vCurrLoginFileName & " " & vStrLine 
                     vConnCount = -100                     
            End Select                       
          end if 
        Loop
        objFile.Close                  
        set objFSO = Nothing 
        set objFile = Nothing  
        makeLogs "vConnCount" & vConnCount
         if vConnCount = 6 then 
           Call makeLogin ()
         else 
            makeLogs "Not all required settings in the file " & vConnCount
         end if 
          

       else 
          makeLogs "Can't find login file " & vCurrLoginFileName   
      end if
end sub 

sub doMDX (vMDXFileName,vMDXLogName)

if len (vSid) < 10 then    
    makeLogs  "Before execute MDX please login ;" 
    exit sub  
end if 

  dim vMDXBody 
  Dim objFSO,objFile
    Set objFSO=CreateObject("Scripting.FileSystemObject")
  Dim vStrLine     

    if objFSO.FileExists(vMDXFileName) then
      set objFile = objFSO.OpenTextFile(vMDXFileName,1)

      vMDXBody = ""
      Do Until objFile.AtEndOfStream
        vStrLine = objFile.ReadLine
        vMDXBody = vMDXBody &  vStrLine
      Loop
      objFile.Close                  
      set objFSO = Nothing 
      set objFile = Nothing  

      strCalcStatus = checkMDXValue (vAPS,vSID,vMDXBody)    
      makeLogs  vMDXLogName & ";" & strCalcStatus  
     else 
          makeLogs "Can't find mdx file " & vMDXFileName        
    end if  
end sub 


sub doCSC (vCSCName )

if len (vSid) < 10 then    
    makeLogs  "Before execute CalcScript  please login ;" 
    exit sub  
end if 
 
  makeLogs  vCSCName & ";" & "Start calculation "  &  chr(13)
  strCalcStatus = execCalcScriptSync(vAPS,vSID,vDB,vCSCName)     
  makeLogs vCSCName   &  strCalcStatus & ";" & "end calculation "

end sub 

sub runScript (vScrptFileName,vCurrPath)
 
		   Dim objFSO,objFile
       Dim vStrLine
       Dim vCurrCommand , vCurrTime , vTimeDuration ,vIsError,vCurrScrptFileName
       vCurrScrptFileName = vCurrPath & "\Commands\" & vScrptFileName 
       makeLogs vCurrScrptFileName
       
        Set objFSO=CreateObject("Scripting.FileSystemObject")
            
		  if objFSO.FileExists(vCurrScrptFileName) then
        set objFile = objFSO.OpenTextFile(vCurrScrptFileName,1)
            dim arrPath 

        
        Do Until objFile.AtEndOfStream
          vStrLine = objFile.ReadLine
          'makeLogs vStrLine
          if (instr(vStrLine,"#")=0 ) then 
            vCurrCommand = split(vStrLine,"=")  
            makeLogs  vStrLine &  ";start" & ";"          
            if Ubound (vCurrCommand)=1 then 
            vCurrTime = Now   
            vIsError = 0        
              Select Case ucase(vCurrCommand(0))
                Case "SLEEP"
                    doSleep vCurrCommand(1) 
                Case "CON"
                    doLogin vCurrPath & "\Commands\" & vCurrCommand(1)                   
                Case "MDX"
                      doMDX vCurrPath & "\mdx\" & vCurrCommand(1) ,vCurrCommand(1)  
                Case "CSC"
                      doCSC vCurrCommand(1)           
                Case Else
                      makeLogs "error " &  vCurrScrptFileName & " " & vStrLine 
                      vIsError = 1
              End Select            
              if  vIsError = 0 then                         
                makeLogs  vStrLine &  ";finished" & ";" & DateDiff("s", vTimeDuration, Now) 
              end if   
            end if 
        end if    
        Loop
        objFile.Close          
        set objFSO = Nothing 
        set objFile = Nothing  
       else 
          makeLogs "Can't find scenario file " & vCurrScrptFileName
      end if
end sub 

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

' main 

 dim strPath
 
 strPath =  getCurrentRunFolder()

 includeFile strPath & "\svStressTool_svc_log.vbs"
 includeFile strPath & "\svStressTool_svc_XML_HTTP.vbs"

 strLogFileName = "CONN"
 dim args
 Set args = Wscript.Arguments

 dim i ,vCurrCommand 
 i = 0 
  For Each arg In args 
 ' writeConole  arg
  i = i + 1
   if i = 1 then 
    vCurrCommand = arg 
   else 
    strLogFileName = arg 
   end if   
  Next

Dim TimeExecution          
    TimeExecution = Now   

   call runScript ( vCurrCommand,strPath) 
    strLogFileName = replace(strLogFileName,"_all.log" ,"_sum.log")      
    makeLogs  "ITERATION RUN ;" & DateDiff("s", TimeExecution, Now)  
 


      


     


  



