<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/DD_DB.asp" -->
<%


dim dbg
' By setting following value to 1 it will show all the interpreted SQL statements in full and execute them
dbg = 0

'response.Write(request.form("totalweight"))
'response.Flush()
'response.End()

Session("ma") = trim(request.form("ma"))
Session("mi") = request.Form("mi")


dim tbc
tbc = replace(request.form("trackbc"),"%"," ")
if len(tbc) = 0 then
	response.Redirect("orders.asp?Customer=" + Session("Customer") + "&Supplier=" + request.form("Supplier") & "&TPName=" & request.form("TPName") & "&FLA=" & Session("FLA") & "&SLA=" & Session("SLA"))
end if

if len(request.form("T1")) < 20 then
	response.Redirect("orders.asp?Customer=" + Session("Customer") + "&Supplier=" + request.form("Supplier") & "&TPName=" & request.form("TPName") & "&FLA=" & Session("FLA") & "&SLA=" & Session("SLA"))	
end if


Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_DD_DB_STRING
Recordset1.Source = "SELECT SUM(WT) as SWT,SUM(PC) AS SPC  FROM DB2ADMIN.ORDERVW2  WHERE DD_SUPPLIER_CODE='" + request.form("Supplier") + "' and DD_PARENT_NO='" + request.Form("trackbc") + "'"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
if dbg = 1 then
  response.Write(Recordset1.Source)
  response.Write("<BR>")
end if

Recordset1_numRows = 0
%>
<%
dim cmdstub

set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = MM_DD_DB_STRING
Command1.CommandType = 1
Command1.CommandTimeout = 30
Command1.Prepared = true
cmdstub = "UPDATE Orders  SET DD_PARENT_NO  WHERE DD_BARCODE_NO ='"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>
<%


dim i
dim updstr 

dim t
t = replace(request.form("T1"),"%"," ")



dim ent
dim mi2
'mi2 = trim(request.form("mi"))
mi2 = trim(replace(Session("mi"),"%"," "))

while len(mi2) < 5 
	mi2 = "0" & mi2
wend

ent = len(t) / 17
while ( i < ent )
	dim bc
	bc = mid(t,(i*18)+1,16)
	if len(trim(bc)) = 16 then
		updstr = "UPDATE Orders  SET DD_PARENT_NO='" + request.Form("trackbc") + "',ORDER_KEY='" & rtrim(request.form("ma")) & mi2 & "' WHERE DD_BARCODE_NO ='" + bc + "' and DD_SUPPLIER_CODE='" + request.form("Supplier") + "'"

		updstr = "UPDATE Orders  SET DD_PARENT_NO='" + request.Form("trackbc") + "',ORDER_KEY='" & trim(Session("ma")) & mi2 & "',PRIMARY_ITEM_IND='S',NO_OF_LABELS=0 WHERE DD_BARCODE_NO ='" + bc + "' and DD_SUPPLIER_CODE='" + request.form("Supplier") + "'"
		Command1.CommandText = updstr
		Command1.CommandTimeout = 0
		
'		response.Write(Command1.CommandText)
'		response.Write("<BR>")
		Command1.Execute()

		updstr = "UPDATE Orders  SET DD_BARCODE_ORIG='" + bc  + "',DD_BARCODE_NO='" + request.form("trackbc") + "' WHERE DD_BARCODE_NO ='" + bc + "' and DD_SUPPLIER_CODE='" + request.form("Supplier") + "'"
		Command1.CommandText = updstr
		Command1.CommandTimeout = 0
		Command1.Execute()

'		response.Write(Command1.CommandText)
'		response.Write("<BR>")
if dbg = 1 then
  response.Write(Command1.CommandText)
  response.Write("<BR>")
end if
	end if
	i = i + 1

wend

Recordset1.Open()
if not Recordset1.Eof then
	dim p ,w
	if isnull(Recordset1.Fields.Item("SPC").Value) then
	    p = 0
	else
		p = Recordset1.Fields.Item("SPC").Value
	end if
	if isnull(Recordset1.Fields.Item("SWT").Value) then
	    w = 0
	else
		w = Recordset1.Fields.Item("SWT").Value
	end if
	
	updstr = "UPDATE Orders SET PRIMARY_ITEM_IND='P',NO_OF_LABELS=1,CPC=" + cstr(p) + ",CWT=" + cstr(w) + " where DD_BARCODE_NO='" + request.Form("trackbc") + "' and DD_SUPPLIER_CODE='" + request.form("Supplier") + "' and DD_BARCODE_NO=DD_BARCODE_ORIG"
	Command1.CommandText = updstr
if dbg = 1 then
  response.Write(Command1.CommandText)
  response.Write("<BR>")
end if
'		response.Write(updstr)
'		response.Write("<BR>")
'		response.Flush()
	Command1.Execute()
'	if err.number <> 0 then
'		response.Write("Error in updating Order <BR><BR>")
'		response.Write(updstr)
'		response.Write("<BR>")
'		response.Flush()
'		response.End()
'	end if
end if

if dbg = 1 then
  response.Flush()
  response.End()
end if
'  response.End()

response.Redirect("orders.asp?Customer=" + Session("Customer") + "&Supplier=" + request.form("Supplier") & "&TPName=" & request.form("TPName") & "&FLA=" & Session("FLA") & "&SLA=" & Session("SLA"))
%>

<body>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>