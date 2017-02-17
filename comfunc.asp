<%

' ###############################################################################################
' #  DDTrack Rewrite Project                                        Copyright OTTO 2012
' #  
' #  This page has been written as part of the new DDTrack project
' #  comfunc.asp
' #  
' #  This routine contains common routines to be used by other ASP files. It also contains code
' #  that runs to check for SQL Injection attacks on any forms or URL strings being passed.
' #   
' #  Change History   
' #  ==============
' #  20/11/2012 Graham Leighton
' #  Initial Build
' ###############################################################################################
%>



<%
Function NZ(ValueIfNotNull, ValueIfNull) 
        NZ = ValueIfNotNull 
        If (IsNull(NZ)) Then NZ = ValueIfNull 
End Function 

Function getRTAG()
	randomize
	getRTAG = CInt(rnd() * 9999)
end function

Function getOrdersStandardCriteria()
	getOrdersStandardCriteria = " TP91_TSTAMP is not null and TP93_TSTAMP is null and TP94_TSTAMP is null and TP9A_TSTAMP is null and TP9B_TSTAMP is null and TP9K_TSTAMP is null and TP9D_TSTAMP is null"
end function



%>


<%

' inline code to check for SQL Injection etc

Dim StandardErrMsg
StandardErrMsg = "An error occurred. Please hit Back to go to the last page and try again"
if request.QueryString <> "" then
    if instr(request.QueryString,"<%") > 0 then
		response.Write(StandardErrMsg)
		response.Flush()
		response.End()
	end if
    if instr(request.QueryString,"JAVASCRIPT") > 0 then
		response.Write(StandardErrMsg)
		response.Flush()
		response.End()
	end if
end if

if request.form <> "" then
    if instr(request.form,"<%") > 0 then
		response.Write(StandardErrMsg)
		response.Flush()
		response.End()
	end if
    if instr(request.form,"JAVASCRIPT") > 0 then
		response.Write(StandardErrMsg)
		response.Flush()
		response.End()
	end if
end if

%>

