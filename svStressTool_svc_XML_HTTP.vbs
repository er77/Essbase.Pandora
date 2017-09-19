' WebServices 

Dim objServerXMLHTTPASync
Dim vCalcCurrScriptName
Dim vPreviousScriptName

function getSVXMLAnswer (vCurrURL,vCurrRequest)
	On Error  Goto 0  'Resume next
	Dim objLocalServerXMLHTTP

	set objLocalServerXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	Dim lResolve,lConnect,lSend,lReceive
	
	objLocalServerXMLHTTP.open "POST", vCurrURL,false
	objLocalServerXMLHTTP.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
	objLocalServerXMLHTTP.setRequestHeader "Content-Length", Len(vCurrRequest)
	objLocalServerXMLHTTP.send vCurrRequest
	
	Dim objDOMDocument
	Set objDOMDocument = CreateObject("MSXML.DOMDocument")
	objDOMDocument.loadXML  Trim(Replace(objLocalServerXMLHTTP.responseXML.xml, vbCrLf, ""))
	
	set objLocalServerXMLHTTP = nothing 	
	
	set getSVXMLAnswer = objDOMDocument 	 
	set objDOMDocument = nothing 	
	
End Function

function getSVXMLAnswerWhaitToEnd (vCurrURL,vCurrRequest)
	On Error  Goto 0  'Resume next
	Dim objSVXMLAnswer	
	set objSVXMLAnswer = CreateObject("MSXML2.ServerXMLHTTP")

	objSVXMLAnswer.setTimeouts 5000, 5000, 15000, 10000000 
	
	objSVXMLAnswer.open "POST", vCurrURL,false
	objSVXMLAnswer.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
	objSVXMLAnswer.setRequestHeader "Content-Length", Len(vCurrRequest)
	objSVXMLAnswer.send vCurrRequest
	
	Dim objDOMDocument
	Set objDOMDocument = CreateObject("MSXML.DOMDocument")
	objDOMDocument.loadXML  Trim(Replace(objSVXMLAnswer.responseXML.xml, vbCrLf, ""))
	
	set objSVXMLAnswer = nothing 	
	
	set getSVXMLAnswerWhaitToEnd = objDOMDocument 	 
	set objDOMDocument = nothing 	
	
End Function

Sub getSVXMLAnswerAsyncCalc (vCurrURL,vCurrRequest,vCurrCalcScriptName,vCurrPreviousScriptName)
	
	On Error  Goto 0  'Resume next
	strHTML = OutStatusCalculation.InnerHTML 
	
	vCalcCurrScriptName = vCurrCalcScriptName
	vPreviousScriptName = vCurrPreviousScriptName
	set objServerXMLHTTPASync = CreateObject("MSXML2.ServerXMLHTTP")  
	objServerXMLHTTPASync.setTimeouts 5000, 5000, 15000, 10000000 
	
	objServerXMLHTTPASync.OnReadyStateChange = GetRef("writeCalcStatus")     
	
	objServerXMLHTTPASync.open "POST", vCurrURL,true
	objServerXMLHTTPASync.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
	objServerXMLHTTPASync.setRequestHeader "Content-Length",  Len(vCurrRequest)
	
	objServerXMLHTTPASync.send vCurrRequest
	
End Sub

function writeCalcStatus  ( )
	
	 
	if (4 <>  objServerXMLHTTPASync.readyState )  then 
		getSleepy
		exit function 
	end if      
	
	dim strHTML 
	strHTML = OutStatusCalculation.InnerHTML  
	Dim objDOMDocument
	Set objDOMDocument = CreateObject("MSXML.DOMDocument")
	objDOMDocument.loadXML  Trim(Replace(objServerXMLHTTPASync.responseXML.xml, vbCrLf, ""))        
	set objServerXMLHTTPASync = nothing 
	
	Dim strStatusCalculation
	strStatusCalculation =  getXMLValue(objDOMDocument,"/vScriptName") 	
	set objDOMDocument = nothing    
	
	if ( Instr(strStatusCalculation,"11") > 0 ) then                
		dim strHTML2,objDivID
		 if (len (vPreviousScriptName) > 0 ) then 
		   Set objDivID =  document.getElementById("btnRun" & vPreviousScriptName  )
		  else 
		  Set objDivID =  document.getElementById("btnRun" & vCalcCurrScriptName  )
		 end if   
		objDivID.innerHTML =  objDivID.innerHTML & "  <div class=""description"">  &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp  Finished : "  & Time &  "</div> "  ' <div class=""item"">  <div class=""content"">
		runStatus.InnerHTML = ""  
		if len (vPreviousScriptName) > 0 then 
		   WriteLog  vCalcCurrScriptName & ";" & " finish"
		else 
		   WriteLog  vConnApp   & "." & vConnDb  & "." & vCalcCurrScriptName & ";" & " finish"
		end if      
	else 
		strHTML = strHTML & "<br>" & Time & "  " & vCalcCurrScriptName  & " calc status <BR> " & strStatusCalculation  
	end if   
	'alert  strHTML
	OutStatusCalculation.InnerHTML  = strHTML  
	getSleepy
    
    call setButtonOnFinish	
	doHttpOnReadyStateChange=0
	 if len (vPreviousScriptName) > 0 then 
	    vGlobalRunnedCSCID=vGlobalRunnedCSCID + 1
       call runCurrentRuleByID (vGlobalRunnedCSCID)
	  
	 end if 
end function 

sub setButtonOnFinish
	runbutton.disabled = false 
    hideCsCbutton.disabled = false
    loadCsCbutton.disabled = false 
	fScheduleFormName.cscRUN.disabled = false 
end sub 

function getXMLNodeValue (vCurrobjDOMDocument,vNodeAddress)
	On Error  Resume Next
	dim objCurrNode
	Set objCurrNode = vCurrobjDOMDocument.selectSingleNode(vNodeAddress)
        getXMLNodeValue = objCurrNode.text 
    set objCurrNode = nothing
	On Error  Goto 0 	
end function

function getXMLtagValue(vCurrobjDOMDocument, XMLTag) 
 On Error  Resume Next
      getXMLtagValue = vCurrobjDOMDocument.getElementsByTagName(XMLTag).Item(0).Text
 On Error  Goto 0 
    if err.number <> 0 then 
      makeLogs "mdx error "
      writeXmlToConsole vCurrobjDOMDocument
    end if        
End Function

function getXMLValue (vCurrobjDOMDocument,vNodeAddress)
	On Error  Resume Next
        getXMLValue = vCurrobjDOMDocument.lastchild.xml 
	On Error  Goto 0 		
end function

function getSID ( vURL,vCurrUserName,vCurrPassword) 
	On Error  Goto 0 
	Dim vRequest
	vRequest = "<req_ConnectToProvider> " _
	& "	 <usr>" & vCurrUserName & "</usr>" _
	& "	 <pwd>" & vCurrPassword & "</pwd>" _
	& "</req_ConnectToProvider>"
	
	Dim objDOMDocument
	Set objDOMDocument = getSVXMLAnswer(vURL,vRequest)
        getSID = getXMLNodeValue(objDOMDocument,"/res_ConnectToProvider/sID") 
	Set objDOMDocument = Nothing	
	if 0 = len(getSID)  Then
		pErrorHandler "getSID error" ,1 
	end if 	
end function

function getSSO ( vURL,vCurrSID ) 
	On Error  Goto 0 
	Dim vRequest
	vRequest = "<req_GetSSOToken> " _
	& "  <sID>" & vCurrSID  &"</sID>" _
	& "</req_GetSSOToken>"
	
	Dim objDOMDocument
	Set objDOMDocument = getSVXMLAnswer(vURL,vRequest)
        getSSO = getXMLNodeValue(objDOMDocument,"/res_GetSSOToken/sso") 
	Set objDOMDocument = Nothing
	if 0 = len(getSSO)  Then
		pErrorHandler "getSSO error" ,1 
	end if 	
end function

function getOpennedApplication ( vCurURL,vCurrSID,vCurrSSO,vCurrEssbaseServer,vCurrAppName ) 
	On Error  Goto 0 
	Dim vRequest
	vRequest = "<req_OpenApplication> " _
					& "   <sID>" & vCurrSID & "</sID> " _
					& "   <srv>" & vCurrEssbaseServer & "</srv> " _
					& "   <app>" & vCurrAppName & "</app>    " _
					& "   <sso>" & vCurrSSO & "</sso> " _
			 & "</req_OpenApplication>"
	'alert vRequest
	Dim objDOMDocument
	Set objDOMDocument = getSVXMLAnswer(vCurURL,vRequest)
        getOpennedApplication = getXMLNodeValue(objDOMDocument,"/res_OpenApplication") 
	Set objDOMDocument = Nothing
	if 0 = len(getOpennedApplication)  Then
		pErrorHandler "getOpennedApplication error" ,1 
	end if 
end function

function getOpennedCube ( vURL,vCurrSID,vCurrSSO,vEssbaseServer,vCurrAppName,vCurrCubeName ) 
	On Error  Goto 0 
	Dim vRequest
	vRequest = "<req_OpenCube> " _
					& "   <sID>" & vCurrSID & "</sID> " _
					& "   <srv>" & vEssbaseServer & "</srv> " _
					& "   <app>" & vCurrAppName & "</app>    " _
					& "   <cube>"& vCurrCubeName & "</cube>    " _				
					& "   <sso>" & vCurrSSO & "</sso> " _
					& "</req_OpenCube>"
	
	Dim objDOMDocument
	Set objDOMDocument = getSVXMLAnswer(vURL,vRequest)
        getOpennedCube = getXMLNodeValue(objDOMDocument,"/res_OpenCube") 
	Set objDOMDocument = Nothing
end function

function getCubeVariables ( vURL,vCurrSID,vCurrAppName,vCurrCubeName  ) 
	On Error  Goto 0 
	Dim vRequest
	vRequest = "<req_GetSubVar> " _
				& "   <sID>" & vCurrSID & "</sID> " _
				& "   <app>" & vCurrAppName & "</app> " _	
				& "   <cube>" & vCurrCubeName & "</cube> " _
				& "   <VarName></VarName>" _    
				& "</req_GetSubVar>"			  
	Dim objDOMDocument
	Set objDOMDocument = getSVXMLAnswer(vURL,vRequest) 
        getCubeVariables =  getXMLValue(objDOMDocument,"/") 
	Set objDOMDocument = Nothing
	if 0 = len(getCubeVariables)  Then
		pErrorHandler "getCubeVariables error" ,1 
	end if 	
end function


function getCubeVariablesList ( vURL,vCurrSID,vCurrAppName,vCurrCubeName  ) 
	On Error  Goto 0 
	Dim vStr,vStr1,vStr2,vArr
	vStr = getCubeVariables( vURL,vCurrSID,vCurrAppName,vCurrCubeName  )   
    vArr = split(vStr,"VarNames>")
    vStr1 = vArr(1)
    vStr1 = replace(vStr1,"</","")
    vArr = split(vStr,"VarVals>")
    vStr2 = vArr(1)
    vStr2 = replace(vStr2,"</","")
    getCubeVariablesList = vStr1 & "#" & vStr2
end function

function getCubeScripts ( vURL,vCurrSID,vCurrAppName,vCurrCubeName  )  
	On Error  Goto 0 
	Dim vRequest
	vRequest = "<req_EnumBusinessRules> " _
			& "   <sID>" & vCurrSID & "</sID> " _
			& "   <app>" & vCurrAppName & "</app> " _	
			& "   <cube>" & vCurrCubeName & "</cube> " _				
			& "</req_EnumBusinessRules>"			  
	Dim objDOMDocument
	Set objDOMDocument = getSVXMLAnswer(vURL,vRequest) 
        getCubeScripts =  getXMLValue(objDOMDocument,"/res_EnumBusinessRules") 
	Set objDOMDocument = Nothing
	if 0 = len(getCubeScripts)  Then
		pErrorHandler "getCubeScripts error" ,1 
	end if 	
end function

function getScript ( vURL,vCurrSID,vCurrCalcScriptName  ) 
	On Error  Goto 0 
	Dim vRequest 
	vRequest = "<req_GetCalcScriptAsString> " _
				& "   <calcScriptName>" & vCurrCalcScriptName & "</calcScriptName>" _
				& "   <type>csc</type> " _
				& "   <sID>" & vCurrSID & "</sID> " _                
			& "</req_GetCalcScriptAsString>"			  
	Dim objDOMDocument
	Set objDOMDocument = getSVXMLAnswer(vURL,vRequest)   
	'alert &objDOMDocument
	jsScriptBody =  getXMLNodeValue(objDOMDocument,"/res_ExecuteCalcScriptAsString/calcScriptContent") 
	'alert jsScriptBody
 	 if (len (jsScriptBody) > 10 ) then 
      jsBeautify    
        jsScriptBody = replace (jsScriptBody," -> ","->")
        jsScriptBody = replace (jsScriptBody,"->    ","->")
        jsScriptBody = replace (jsScriptBody,"->    ","->")
        jsScriptBody = replace (jsScriptBody,"->    ","->")
        jsScriptBody = replace (jsScriptBody,"->    ","->")
        jsScriptBody = replace (jsScriptBody,"->		","->") 
        jsScriptBody = replace (jsScriptBody, "> 0 ) ","> 0 )" & chr(13) & chr(9)& chr(9)  )
        jsScriptBody = replace (jsScriptBody, "= 0 ) ","> 0 )" & chr(13) & chr(9)& chr(9)  )
        jsScriptBody = replace (jsScriptBody, "/*13.9.9*/","/*13.9.9*/" & chr(13) & chr(9)& chr(9)  )
	  end if 	
        getScript = jsScriptBody
	Set objDOMDocument = Nothing
	if 0 = len(getScript)  Then
		pErrorHandler "getScript error" ,1 
	end if        
end function

function execCalcScriptSync  (vCurrAps,vCurrSID,vCurrDB,vCurrCalcScript)
	Dim strHTML,objDivID 
	
	if ( instr(Ucase(vCurrCalcScript),Ucase("Default")) =  0  ) then
    strHTML = "<req_LaunchBusinessRule> " _
						& "   <sID>" & vCurrSID & "</sID> " _	                         
						& "   <cube>" & vCurrDB & "</cube> " _					
						& "   <rule>" & vCurrCalcScript & "</rule> " _					
				& "</req_LaunchBusinessRule>"	
    dim XmlResult						   
	 set XmlResult = getSVXMLAnswerWhaitToEnd(vCurrAps,strHTML)                  
    
	Dim objDOMDocument
	Set objDOMDocument = CreateObject("MSXML.DOMDocument")
	objDOMDocument.loadXML  Trim(Replace(XmlResult.xml, vbCrLf, ""))        
    set XmlResult = Nothing
	
	Dim strStatusCalculation
	strStatusCalculation =  getXMLValue(objDOMDocument,"/vScriptName") 	
	set objDOMDocument = nothing  
    end if  

    if ( instr(Ucase(vCurrCalcScript),Ucase("Default")) =  0  ) then
		execCalcScriptSync=""
		if ( Instr(strStatusCalculation,"11") = 0 ) then   
			execCalcScriptSync = strStatusCalculation            		 
		end if 		
	Else 
	  execCalcScriptSync=" 'Default' is restricted calcualtion "	  
	end if  	
	set strStatusCalculation = nothing  
END function 

function getMdxXmlRaw  (vCurrAps,vCurrSID,vCurrMDXBody)
	Dim strHTML 	 
    strHTML = "<req_ExecuteQuery> " _
						& "   <sID>" & vCurrSID & "</sID> " _	                         
						& "   <preferences> " _	
						& "         <row_suppression zero=""0"" invalid=""0"" missing=""0"" underscore=""0"" noaccess=""0""/> " _	
						& "         <celltext val=""1""/> " _	
						& "         <zoomin ancestor=""top"" mode=""children""/>" _	
						& "         <navigate withData=""1""/>" _	
						& "         <includeSelection val=""1""/>" _	
						& "         <repeatMemberLabels val=""1""/>" _	
						& "         <withinSelectedGroup val=""0""/>" _	
						& "         <removeUnSelectedGroup val=""0""/>" _	
						& "         <col_suppression zero=""0"" invalid=""0"" missing=""0"" underscore=""0"" noaccess=""1""/>" _	
						& "         <block_suppression missing=""0""/>" _	
						& "         <includeDescriptionInLabel val=""2""/>" _	
						& "         <missingLabelText val=""""/>" _	
						& "         <noAccessText val=""'-""/>" _	
						& "         <aliasTableName val=""none""/>" _	
						& "         <essIndent val=""1""/>" _	
						& "         <formulaRetention retrieve=""1"" zoomin_out=""1"" keep_removeonly=""1"" fill=""1"" comments=""1""/>" _	
					    & "         <sliceLimitation rows=""1048576"" cols=""16384""/>" _	
				        & "         </preferences> "_				
						& "   <mdx>" & vCurrMDXBody & "</mdx> " _					
				& "</req_ExecuteQuery>"	
    dim XmlResult						   
	 set getMdxXmlRaw = getSVXMLAnswerWhaitToEnd(vCurrAps,strHTML)                  
    'writeXmlToConsole XmlResult         		 		 
	' getMdxXmlRaw	
end function  

function getMDXValue  (vCurrAps,vCurrSID,vCurrMDXBody)
    dim XmlResult		 		   
	  set XmlResult = getMdxXmlRaw(vCurrAps,vCurrSID,vCurrMDXBody)                
    'writeXmlToConsole XmlResult 

	Dim objDOMDocument
	Set objDOMDocument = CreateObject("MSXML.DOMDocument")
	objDOMDocument.loadXML  Trim(Replace(XmlResult.xml, vbCrLf, ""))        
    set XmlResult = Nothing
	
	Dim strStatusCalculation
	getMDXValue =  getXMLtagValue(objDOMDocument,"vals") 	
		if len (getMDXValue) < 2 then 
			getMDXValue ="#err: " &  getXMLNodeValue(objDOMDocument,"/")
		end if  	
	set objDOMDocument = nothing 
	 
end function 

function checkMDXValue  (vCurrAps,vCurrSID,vCurrMDXBody)
    dim vMDXResult						   
	   vMDXResult = getMDXValue(vCurrAps,vCurrSID,vCurrMDXBody)                     

	checkMDXValue = "" 
    if len(vMDXResult) >1 then 
	  if instr(vMDXResult,"#err") > 0 then 
	    checkMDXValue = vMDXResult 
	   else 
	   checkMDXValue = "1"
	 end if 

	end if  
	 	
end function 