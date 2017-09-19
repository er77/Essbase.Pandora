Sub doSleep( vTime  )
  WScript.Sleep vTime * 1000
end Sub

Sub getSleepy
	 doSleep 0.01 
End Sub 

sub writeConole  (LogMessage )
  Wscript.Echo Chr(13) & getRightTime() & ";" & LogMessage
end sub

sub writeXmlToConsole   (vCurrXMLFile ) 
for each x in vCurrXMLFile.documentElement.childNodes
   Wscript.Echo Chr(13) & (x.nodename) & ": " & x.text
   Wscript.Echo Chr(13) 
next
end Sub 
 

function getRightTime 
 Dim vStrTime 
		 vStrTime  =   Year (Now)

		  if ( Month (Now) < 10  ) then 
		   vStrTime = vStrTime &"0" &  Month (Now)
		   else 
		   vStrTime = vStrTime &  Month (Now)
		  end if 

		  if ( Day (Now) < 10  ) then 
		   vStrTime = vStrTime &"0" &  Day (Now)
		   else 
		   vStrTime = vStrTime &  Day (Now)
		  end if 
           vStrTime = vStrTime & " " & Time 
   getRightTime = vStrTime

end function 

 
sub WriteFileLog (vCurrFileName,vCurrLogMessage )
        Dim objFSO,objFile,vStrFileName
        Set objFSO=CreateObject("Scripting.FileSystemObject")        

		 if objFSO.FileExists(vCurrFileName) then 
		    set objFile =  objFSO.OpenTextFile(vCurrFileName, 8, True)', TristateTrue ) 
		 else 
		    set objFile =  objFSO.CreateTextFile(vCurrFileName,true) 	 
		 end if 

         dim LogMessage
		 LogMessage  =  getRightTime() & ";" &  vCurrLogMessage & ";"

          'alert LogMessage   
           objFile.WriteLine LogMessage  
         
            objFile.Close
            set objFSO = Nothing 
            set objFile = Nothing 
End sub
