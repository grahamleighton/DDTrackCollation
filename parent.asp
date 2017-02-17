<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/DD_DB.asp" -->
<%
Dim Recordset1__MM_SUPPCODE
Recordset1__MM_SUPPCODE = "1"
If (request.Form("Supplier") <> "") Then 
  Recordset1__MM_SUPPCODE = request.Form("Supplier")
End If
%>
<%
Dim Recordset1__MM_MASTER
Recordset1__MM_MASTER = "1"
If (request.form("trackbc") <> "") Then 
  Recordset1__MM_MASTER = request.form("trackbc")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_DD_DB_STRING
Recordset1.Source = "SELECT SUM(WT) as SWT,SUM(PC) AS SPC  FROM DB2ADMIN.ORDERVW2  WHERE DD_SUPPLIER_CODE='" + request.form("Supplier") + "' and DD_PARENT_NO='" + request.Form("trackbc") + "'"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1

Recordset1_numRows = 0
%>
<%
dim cmdstub

set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = MM_DD_DB_STRING
cmdstub = "UPDATE Orders  SET DD_PARENT_NO  WHERE DD_BARCODE_NO ='"
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true

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
t = request.Form("T1")



dim ent

ent = len(t) / 17
'response.Write("T1=")
'response.Write(t)
'response.Write("<BR>")

'response.Write("len(T1)=")
'response.Write(len(t))
'response.Write("<BR>")

'response.write("Entries = ")
'response.Write(ent)
'response.Write("<BR>")

'response.End()

while ( i < ent )
	dim bc
	bc = mid(t,(i*18)+1,16)
	if len(trim(bc)) = 16 then
		updstr = "UPDATE Orders  SET DD_PARENT_NO='" + request.Form("trackbc") + "' WHERE DD_BARCODE_NO ='" + bc + "' and DD_SUPPLIER_CODE='" + request.form("Supplier") + "'"
		Command1.CommandText = updstr
'		response.Write(Command1.CommandText)
'		response.Write("<BR>")
'		response.Flush()
		Command1.CommandTimeout = 0
		Command1.Execute()
	end if
	i = i + 1

wend
'response.Flush()
'response.Write("Completed")
Recordset1.Open()
if not Recordset1.Eof then
	updstr = "UPDATE Orders SET CPC=" + cstr((Recordset1.Fields.Item("SPC").Value)) + ",CWT=" + cstr((Recordset1.Fields.Item("SWT").Value)) + " where DD_BARCODE_NO='" + request.Form("trackbc") + "' and DD_SUPPLIER_CODE='" + request.form("Supplier") + "'"
	Command1.CommandText = updstr
'		response.Write(updstr)
'		response.Write("<BR>")
'		response.Flush()
	Command1.Execute()
end if

response.Redirect("orders.asp?Customer=" + Session("Customer") + "&Supplier=" + request.form("Supplier"))
%>

<body>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
